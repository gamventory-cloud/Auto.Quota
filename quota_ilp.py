"""
╔══════════════════════════════════════════════════════════════════════════╗
║  파일명 : quota_ilp.py                                                    ║
║  위치   : 리포지토리 최상단  (pages/ 폴더 안이 아님!)                       ║
║  필요   : pip install ortools                                             ║
╚══════════════════════════════════════════════════════════════════════════╝

quota_ilp.py — 쿼터 할당을 정수계획법(ILP)으로 정확히 푸는 솔버

v2.0 변경점 : 추가 쿼터를 "상한"이 아니라 "목표"로 처리
------------------------------------------------------------------
v1.0 은 추가 쿼터를 상한(≤ C)으로만 걸었다. 초과는 막았지만 부족은
전혀 페널티가 없어서, 추가 쿼터가 절반만 채워져도 솔버는 만족했다.

v2.0 은 추가 쿼터도 메인과 똑같이 **등식 + 부족변수**로 만든다.

    Σ n_p + s = C        (s ≥ 0)      ⟺      Σ n_p ≤ C

제약의 실행 가능 영역은 완전히 동일하다. 달라지는 것은 s 가 목적함수에
들어가서 "부족을 줄이려고 노력한다"는 점뿐이다. 따라서 이 변경으로
INFEASIBLE 이 나는 일은 없고, 달성 불가한 목표는 부족량으로 보고된다.

  ex_as_target=False → v1.0 과 동일 동작 (상한 전용)
  ex_as_target=True  → 추가 쿼터도 100% 채우려 시도

반드시 알아야 하는 산술 항등식
------------------------------------------------------------------
추가 쿼터 그룹에서 응답자 한 명이 정확히 키 하나를 갖는 경우
(조합형, 또는 무응답 없는 단일응답 변수):

    Σ_k 추가달성_k  =  선택 인원 총합

메인을 100% 채우면 우변이 target_total 로 고정되므로,

    Σ_k 추가목표_k  <  target_total  →  초과 금지 하에서 메인 100% 불가
    Σ_k 추가목표_k  >  target_total  →  그 차이만큼 추가 부족이 반드시 발생
    Σ_k 추가목표_k  == target_total  →  이때만 양쪽 100% 가능 (필요조건)

preflight_targets() 가 실행 전에 이 산술을 검사해 정확한 숫자로 알려준다.
솔버를 돌리기 전에 반드시 호출할 것.

핵심 기법 : 프로파일 집약 (v1.0 과 동일)
------------------------------------------------------------------
쿼터 제약은 "각 키를 몇 명 뽑았는가"에만 의존하므로, 동일한
(메인 키, 추가 키 조합) 을 가진 응답자는 서로 완전히 교환 가능하다.
같은 서명끼리 묶어 정수변수 하나로 집약하면 5만 행이 수백~수천 변수가 된다.

수학 모형
------------------------------------------------------------------
  변수   n_p ∈ {0..avail_p}              프로파일 p 에서 뽑을 인원
         sM_k ∈ {0..T_k}                 메인 셀 k 부족분
         sE_(j,k) ∈ {0..C_jk}            추가 쿼터 (j,k) 부족분   ← v2.0

  제약   Σ_{p ∈ cell k} n_p + sM_k    = T_k
         Σ_{p ∋ (j,k)}  n_p + sE_jk   = C_jk        (초과 금지)

  목적   사전식(lexicographic) 다단계
         1) min Σ sM                    메인 부족 최소화 (= 통과 인원 최대)
         2) min Σ_j w_j · Σ_k sE_jk     추가 부족 최소화            ← v2.0
         3) min Σ scarcity_k · sM_k     희소 셀 우선 채우기
         4) min max(모든 부족변수)       남은 부족 고르게 분산

  뒤 단계는 앞 단계 최적값을 등식으로 고정한 뒤 최적화하므로 앞 단계를
  절대 훼손하지 않는다. 2단계가 3단계보다 앞이라는 점이 중요하다.
"""

from __future__ import annotations

import collections
import math
import time
from dataclasses import dataclass, field

try:
    from ortools.sat.python import cp_model
except ImportError as e:  # pragma: no cover
    raise ImportError("OR-Tools가 필요합니다.  pip install ortools") from e


MODULE_ROLE = "quota_ilp"
__version__ = "2.0"


# ==============================================================================
# 결과 컨테이너
# ==============================================================================
@dataclass
class Diagnosis:
    """미달 원인 진단."""
    # 메인 쿼터
    main_short: dict = field(default_factory=dict)       # 셀 -> 부족 인원
    main_avail: dict = field(default_factory=dict)       # 셀 -> 데이터 보유 인원
    main_reason: dict = field(default_factory=dict)      # 셀 -> 사유
    # 추가 쿼터 (그룹별 dict 리스트)
    ex_short: list = field(default_factory=list)
    ex_over: list = field(default_factory=list)      # 초과 (소프트 쿼터에서만)
    ex_avail: list = field(default_factory=list)
    ex_reason: list = field(default_factory=list)
    # 병목 / 민감도
    binding: list = field(default_factory=list)          # 한도까지 꽉 찬 추가 쿼터
    group_relax_gain: dict = field(default_factory=dict) # 그룹 초과금지 해제 시 이득
    value_relax_gain: dict = field(default_factory=dict) # 특정 한도 +1 시 이득
    # 산술 정합성 (preflight 결과를 그대로 담아둔다)
    arithmetic: list = field(default_factory=list)


@dataclass
class IlpSolution:
    status: str
    proven_optimal: bool
    selected: list
    total: int
    target_total: int
    main_actual: dict
    ex_actual: list
    diagnosis: Diagnosis
    n_profiles: int
    n_rows_considered: int
    solve_sec: float
    notes: list = field(default_factory=list)
    # v2.0 추가
    ex_as_target: bool = False
    main_short_total: int = 0
    ex_short_total: list = field(default_factory=list)   # 그룹별 부족 합계
    ex_over_total: list = field(default_factory=list)    # 그룹별 초과 합계
    stage_values: list = field(default_factory=list)     # [(단계명, 최적값), ...]
    all_satisfied: bool = False                          # 메인+추가 전부 100%
    main_hard: bool = False                              # 하드 쿼터로 풀렸는가
    ex_overflow: bool = False                            # 초과 허용 여부


# ==============================================================================
# 0. 실행 전 산술 정합성 검사
# ==============================================================================
def preflight_targets(m_keys, ex_keys_list, main_map, ex_maps,
                      ex_as_target=True, unlisted="free", ex_overflow=False):
    """
    솔버를 돌리기 전에 "애초에 달성 가능한 목표인가"를 검사한다.

    반환: [{'level', 'group', 'kind', 'msg', 'numbers'}, ...]
      level : 'error' = 구조적으로 달성 불가
              'warn'  = 부족이 확정적으로 발생
              'info'  = 참고
    """
    out = []
    target_total = sum(main_map.values())
    n_rows = len(m_keys)

    # --- 선택 가능한 행 (메인 목표가 0보다 큰 셀) ---
    elig = [i for i in range(n_rows) if main_map.get(m_keys[i], 0) > 0]
    m_avail = collections.Counter(m_keys[i] for i in elig)

    # 1) 메인 쿼터 유령 셀 / 물리적 부족
    ghosts = [k for k, t in main_map.items() if t > 0 and m_avail.get(k, 0) == 0]
    if ghosts:
        out.append({
            'level': 'error', 'group': None, 'kind': 'main_ghost',
            'msg': (f"메인 쿼터 {len(ghosts)}개 셀에 해당하는 응답자가 데이터에 "
                    f"한 명도 없습니다. 목표 "
                    f"{sum(main_map[k] for k in ghosts):,}명은 달성 불가입니다."),
            'numbers': {'cells': ghosts[:20], 'n_cells': len(ghosts),
                        'lost': sum(main_map[k] for k in ghosts)},
        })
    phys = {k: t - m_avail.get(k, 0) for k, t in main_map.items()
            if m_avail.get(k, 0) < t}
    if phys:
        out.append({
            'level': 'warn', 'group': None, 'kind': 'main_phys',
            'msg': (f"메인 쿼터 {len(phys)}개 셀이 보유 인원보다 목표가 큽니다. "
                    f"최소 {sum(phys.values()):,}명 부족이 확정입니다."),
            'numbers': {'shortfall': sum(phys.values()), 'cells': dict(list(phys.items())[:20])},
        })

    total_avail = len(elig)
    if total_avail < target_total:
        out.append({
            'level': 'error', 'group': None, 'kind': 'main_total',
            'msg': (f"선택 가능한 응답자가 {total_avail:,}명뿐인데 메인 목표 합계는 "
                    f"{target_total:,}명입니다."),
            'numbers': {'avail': total_avail, 'target': target_total},
        })

    if not ex_as_target:
        return out

    # 2) 그룹별 산술 검사
    for j, e_map in enumerate(ex_maps):
        if not e_map:
            continue
        cap_sum = sum(e_map.values())

        # 응답자별 '제약이 걸린 키' 개수
        counts = []
        na_rows = 0
        unlisted_rows = 0
        for i in elig:
            ks = ex_keys_list[j][i]
            known = [k for k in ks if k in e_map]
            if not ks:
                na_rows += 1
            if len(known) < len(ks):
                unlisted_rows += 1
            counts.append(len(known))

        if not counts:
            continue

        # 목표 인원만큼 뽑을 때 Σ 달성 가능 범위 (메인 제약 무시한 느슨한 경계)
        sc = sorted(counts)
        take = min(target_total, len(sc))
        lo = sum(sc[:take])
        hi = sum(sc[-take:]) if take else 0

        if cap_sum < lo and ex_overflow:
            out.append({
                'level': 'warn', 'group': j, 'kind': 'group_sum_low_soft',
                'msg': (f"추가 쿼터 목표 합계가 {cap_sum:,}명인데 메인 목표를 채우려면 "
                        f"최소 {lo:,}명이 계상됩니다. 초과를 허용했으므로 메인은 "
                        f"100% 달성되고, 추가 쿼터가 최소 {lo - cap_sum:,}명 "
                        f"초과합니다."),
                'numbers': {'cap_sum': cap_sum, 'need_min': lo,
                            'overflow': lo - cap_sum, 'target_total': target_total},
            })
        elif cap_sum < lo:
            out.append({
                'level': 'error', 'group': j, 'kind': 'group_sum_low',
                'msg': (f"추가 쿼터 목표 합계가 {cap_sum:,}명인데, 메인 목표 "
                        f"{target_total:,}명을 채우려면 이 그룹에서 최소 {lo:,}명이 "
                        f"계상되어야 합니다. 초과가 금지되어 있으므로 "
                        f"**메인 100% 달성이 불가능**합니다. "
                        f"목표 합계를 {lo:,}명 이상으로 올리세요."),
                'numbers': {'cap_sum': cap_sum, 'need_min': lo, 'target_total': target_total},
            })
        elif cap_sum > hi:
            out.append({
                'level': 'warn', 'group': j, 'kind': 'group_sum_high',
                'msg': (f"추가 쿼터 목표 합계가 {cap_sum:,}명인데 최대 {hi:,}명까지만 "
                        f"계상될 수 있습니다. 최소 {cap_sum - hi:,}명 부족이 확정입니다. "
                        f"(메인 목표 합계 {target_total:,}명)"),
                'numbers': {'cap_sum': cap_sum, 'reach_max': hi,
                            'shortfall': cap_sum - hi, 'target_total': target_total},
            })
        else:
            out.append({
                'level': 'info', 'group': j, 'kind': 'group_sum_ok',
                'msg': (f"추가 쿼터 목표 합계 {cap_sum:,}명 — 달성 가능 범위 "
                        f"{lo:,}~{hi:,}명 안에 있습니다."),
                'numbers': {'cap_sum': cap_sum, 'lo': lo, 'hi': hi},
            })

        if na_rows:
            out.append({
                'level': 'warn', 'group': j, 'kind': 'na_rows',
                'msg': (f"이 그룹의 변수가 무응답인 응답자가 {na_rows:,}명입니다. "
                        f"이들을 뽑으면 추가 쿼터에 1도 계상되지 않아, 그만큼 "
                        f"목표 합계를 채울 수 없습니다."),
                'numbers': {'na_rows': na_rows},
            })

        if unlisted_rows:
            out.append({
                'level': 'info' if unlisted == 'free' else 'warn',
                'group': j, 'kind': 'unlisted',
                'msg': (f"목표 목록에 없는 값을 가진 응답자가 {unlisted_rows:,}명입니다. "
                        + ("제약 없이 자유롭게 선택됩니다 (unlisted='free')."
                           if unlisted == 'free'
                           else "선택 대상에서 제외됩니다 (unlisted='forbid').")),
                'numbers': {'unlisted_rows': unlisted_rows},
            })

        # 3) 키별 물리적 부족 / 유령 키
        avail_k = collections.Counter()
        for i in elig:
            for k in ex_keys_list[j][i]:
                if k in e_map:
                    avail_k[k] += 1
        g_ghost = [k for k, c in e_map.items() if c > 0 and avail_k.get(k, 0) == 0]
        if g_ghost:
            out.append({
                'level': 'error', 'group': j, 'kind': 'ex_ghost',
                'msg': (f"추가 쿼터 {len(g_ghost)}개 항목에 해당하는 응답자가 "
                        f"데이터에 없습니다. 목표 "
                        f"{sum(e_map[k] for k in g_ghost):,}명은 달성 불가입니다."),
                'numbers': {'keys': [str(k) for k in g_ghost[:20]],
                            'lost': sum(e_map[k] for k in g_ghost)},
            })
        g_phys = {k: c - avail_k.get(k, 0) for k, c in e_map.items()
                  if avail_k.get(k, 0) < c}
        if g_phys:
            out.append({
                'level': 'warn', 'group': j, 'kind': 'ex_phys',
                'msg': (f"추가 쿼터 {len(g_phys)}개 항목이 보유 인원보다 목표가 큽니다. "
                        f"최소 {sum(g_phys.values()):,}명 부족이 확정입니다."),
                'numbers': {'shortfall': sum(g_phys.values()),
                            'keys': {str(k): v for k, v in list(g_phys.items())[:20]}},
            })

    return out


# ==============================================================================
# 1. 프로파일 집약
# ==============================================================================
def build_profiles(m_keys, ex_keys_list, main_map, ex_maps,
                   ex_as_target=False, unlisted="free"):
    """
    응답자를 동일 서명 단위로 묶는다.

    서명 = (메인 키, 그룹별 '제약이 걸린 키'를 정렬한 튜플)

    - 메인 목표가 없거나 0인 셀의 응답자는 애초에 뽑을 수 없으므로 제외
    - 목표 0 인 추가 키는 **금지**를 의미하므로 서명에 남긴다 (v1.0 은 제거해서
      목표 0 이 '무제한'으로 뒤집히는 버그가 있었다)
    - 목표 목록에 아예 없는 키는 unlisted 규칙에 따른다
        'free'   : 제약 없음 → 서명에서 제거
        'forbid' : 선택 금지 → 해당 행을 제외

    반환: [(서명, [행 위치, ...]), ...]
    """
    n_rows = len(m_keys)
    n_groups = len(ex_maps)
    buckets = collections.defaultdict(list)

    for i in range(n_rows):
        mk = m_keys[i]
        if main_map.get(mk, 0) <= 0:
            continue

        sig_ex = []
        drop = False
        for j in range(n_groups):
            e_map = ex_maps[j]
            if not e_map:
                sig_ex.append(())
                continue
            ks = ex_keys_list[j][i]
            if unlisted == "forbid" and any(k not in e_map for k in ks):
                drop = True
                break
            if ex_as_target:
                # 목표 모드: 목록에 있는 키는 0이든 양수든 모두 제약 대상
                constrained = {k for k in ks if k in e_map}
            else:
                # 상한 모드: 0 은 금지, 목록에 없으면 제약 없음
                constrained = {k for k in ks if k in e_map and e_map[k] >= 0}
            sig_ex.append(tuple(sorted(constrained, key=repr)))
        if drop:
            continue

        buckets[(mk, tuple(sig_ex))].append(i)

    return list(buckets.items())


def scarcity_weights(profiles, main_map, scale=1000, floor=0.01):
    """
    메인 셀별 희소성 가중치. 보유/목표 비율이 낮을수록 크게 준다.
    3단계 목적함수에만 쓰이므로 총 인원을 깎지 않는 순수 타이브레이크다.
    """
    avail = collections.Counter()
    for (mk, _), rows in profiles:
        avail[mk] += len(rows)
    w = {}
    for k, tgt in main_map.items():
        if tgt <= 0:
            continue
        r = avail.get(k, 0) / tgt
        w[k] = int(round(scale / max(r, floor)))
    return w


# ==============================================================================
# 2. 모형 구성 및 해 구하기
# ==============================================================================
def _squared_dev_terms(model, pairs, tag=""):
    """
    Σ w_k · 편차_k²  항을 만든다.  w_k = max(1, round(1e6 / 목표²))

    최소최대(min-max)만 쓰면, 물리적으로 불가피한 큰 편차 하나가 최댓값을
    포화시켜 나머지를 고르게 나눌 동기가 사라진다. 제곱 편차를 뒤 단계에
    두면 남은 편차가 여러 항목에 자연스럽게 분산된다.
      예) 84를 한 항목에 몰기(84²=7056) vs 세 항목에 28씩(3×784=2352)

    pairs : [(편차 IntVar 후보 리스트, 목표수, 편차 상한), ...]
            상한은 호출부가 명시한다. ortools 9.15 에서는 변수 도메인을
            Proto() 로 들여다보면 메모리 오류가 발생하므로 절대 쓰지 않는다.
    """
    terms = []
    for idx, (vars_, cap, ub) in enumerate(pairs):
        cap = max(1, int(cap))
        ub = int(ub)
        if not vars_ or ub <= 0:
            continue
        d = model.NewIntVar(0, ub, f"dev{tag}{idx}")
        model.Add(d == sum(vars_))
        sq = model.NewIntVar(0, ub * ub, f"sq{tag}{idx}")
        model.AddMultiplicationEquality(sq, [d, d])
        w = max(1, int(round(1_000_000 / (cap * cap))))
        terms.append(w * sq)
    return terms


def _solve_core(profiles, main_map, ex_maps, weights=None,
                ex_as_target=False, ex_weights=None,
                main_hard=False, ex_overflow=False, overflow_weight=1,
                ex_tol_abs=0, ex_tol_pct=0.0, ex_tol_unlimited=False,
                skip_groups=frozenset(), cap_bonus=None,
                priority=False, balance=False, balance_relative=True,
                min_fill=0.0, time_limit=30.0, workers=8, log=False):
    """
    집약된 모형을 사전식 다단계로 푼다. 내부용.

    main_hard  : True 면 메인 부족변수 상한을 0 으로 묶어 **정확히 목표대로** 뽑는다.
                 물리적으로 불가능하면 INFEASIBLE 이 되므로 호출부가 자동 완화한다.
    ex_overflow: True 면 추가 쿼터가 목표에서 벗어나는 것을 허용한다.
                 총 선정 인원은 메인 쿼터가 정하므로 바뀌지 않고, 개별 항목만
                 목표 위아래로 흔들린다. 예: 50/50/50/50 → 55/45/47/53 (합 200 유지)
    ex_tol_abs / ex_tol_pct : 항목별 허용 편차 한계. band = max(abs, ceil(목표×pct))
                 부족과 초과에 **대칭으로** 걸린다. 둘 다 0 이면 편차 0 (하드).
    ex_tol_unlimited : True 면 편차 한계 없이 최소화만 한다.

    반환: (ok, status, solver, n_vars, short_main, short_ex, over_ex, stage_values)
      ok : True / None(실패) / "GUARD_INFEASIBLE"
    """
    model = cp_model.CpModel()

    # --- 변수: 프로파일별 선택 인원 ---
    n_vars = [model.NewIntVar(0, len(rows), f"n{idx}")
              for idx, (_sig, rows) in enumerate(profiles)]

    # --- 메인 쿼터: 등식 + 부족 슬랙 ---
    by_main = collections.defaultdict(list)
    for idx, ((mk, _), _rows) in enumerate(profiles):
        by_main[mk].append(idx)

    short_main = {}
    for k, tgt in main_map.items():
        # main_hard 면 부족을 0 으로 묶는다 = 하드 쿼터 (정확히 tgt 명)
        s = model.NewIntVar(0, 0 if main_hard else tgt, f"sM{len(short_main)}")
        short_main[k] = s
        model.Add(sum(n_vars[i] for i in by_main.get(k, [])) + s == tgt)

    # --- 추가 쿼터 ---
    by_ex = collections.defaultdict(list)
    for idx, ((_mk, sig_ex), _rows) in enumerate(profiles):
        for j, keys in enumerate(sig_ex):
            if j in skip_groups:
                continue
            for k in keys:
                by_ex[(j, k)].append(idx)

    short_ex = {}
    over_ex = {}
    for j, e_map in enumerate(ex_maps):
        if not e_map or j in skip_groups:
            continue
        for k, cap in e_map.items():
            eff = cap + (cap_bonus or {}).get((j, k), 0)
            members = by_ex.get((j, k), [])
            if eff <= 0:
                # 목표 0 = 금지. 초과 허용 여부와 무관하게 강제 0 으로 묶는다.
                for i in members:
                    model.Add(n_vars[i] == 0)
                continue
            if not members:
                # 데이터에 없는 키. 목표 모드에서는 전량 부족으로 계상한다.
                if ex_as_target:
                    s = model.NewIntVar(eff, eff, f"sE_ghost{len(short_ex)}")
                    short_ex[(j, k)] = s
                continue

            avail_k = sum(len(profiles[i][1]) for i in members)
            if ex_as_target or ex_overflow:
                # 항목별 허용 편차 밴드 (부족·초과 대칭)
                if not ex_overflow:
                    band_s, band_o = eff, 0        # 부족만 허용, 초과 금지
                elif ex_tol_unlimited:
                    band_s, band_o = eff, max(0, avail_k - eff)
                else:
                    band = max(int(ex_tol_abs),
                               int(math.ceil(eff * float(ex_tol_pct))))
                    band_s = min(eff, band)
                    band_o = min(max(0, avail_k - eff), band)
                s = model.NewIntVar(0, band_s, f"sE{len(short_ex)}")
                short_ex[(j, k)] = s
                o = model.NewIntVar(0, band_o, f"oE{len(over_ex)}")
                over_ex[(j, k)] = o
                # Σ + 부족 − 초과 = 목표
                model.Add(sum(n_vars[i] for i in members) + s - o == eff)
            else:
                model.Add(sum(n_vars[i] for i in members) <= eff)

    total_avail = sum(len(rows) for _sig, rows in profiles)

    # --- 목적함수 단계 구성 ---
    w = weights or {}
    stages = [("main", sum(w.get(k, 1) * s for k, s in short_main.items()))]

    if (ex_as_target or ex_overflow) and short_ex:
        ew = ex_weights or {}
        terms = []
        for (j, k), sv in short_ex.items():
            wj = int(ew.get(j, 1))
            terms.append(wj * sv)
            ov = over_ex.get((j, k))
            if ov is not None:
                terms.append(wj * int(overflow_weight) * ov)
        stages.append(("ex", sum(terms)))

    sw = scarcity_weights(profiles, main_map) if priority else None
    if priority:
        stages.append(("scarcity", sum(sw.get(k, 1) * s
                                       for k, s in short_main.items())))

    if balance:
        # 부족을 어느 항목이 감당할지 고르게 분산한다.
        #
        # balance_relative=True (기본) : '목표 대비 부족률' 기준으로 분산한다.
        #   목표 1000인 셀의 50명 부족(5%)과 목표 100인 셀의 50명 부족(50%)을
        #   똑같이 취급하면 작은 셀이 통째로 비어버린다. 그래서 부족량에
        #   BIG/목표 를 곱해 비율로 환산한 뒤 그 최댓값을 최소화한다.
        # balance_relative=False : 절대 인원수 기준 (v1.0 동작)
        BIG = 10000
        scaled = []
        for k, s in short_main.items():
            t = max(1, main_map[k])
            scaled.append((BIG // t if balance_relative else 1) * s)
        for (j, k), s in short_ex.items():
            c = max(1, ex_maps[j].get(k, 1))
            scaled.append((BIG // c if balance_relative else 1) * s)
        for (j, k), o in over_ex.items():
            c = max(1, ex_maps[j].get(k, 1))
            scaled.append((BIG // c if balance_relative else 1) * o)
        if scaled:
            ub = BIG if balance_relative else max(
                [t for t in main_map.values()] +
                [c for m in ex_maps if m for c in m.values()] + [1])
            worst = model.NewIntVar(0, max(1, ub), "worst")
            model.AddMaxEquality(worst, scaled)
            stages.append(("balance", worst))

        # 최소최대는 불가피한 큰 편차 하나에 포화되므로, 제곱 편차로 마무리한다
        pairs = []
        for k, sv in short_main.items():
            t = main_map[k]
            pairs.append(([sv], t, 0 if main_hard else t))
        for (j, k), sv in short_ex.items():
            cap = ex_maps[j].get(k, 1)
            ov = over_ex.get((j, k))
            vs = [sv] + ([ov] if ov is not None else [])
            # 편차 상한 : 부족은 최대 cap, 초과는 선정 가능 인원까지
            pairs.append((vs, cap, cap + total_avail))
        sq_terms = _squared_dev_terms(model, pairs, tag="b")
        if sq_terms:
            stages.append(("spread", sum(sq_terms)))

    def _run():
        s = cp_model.CpSolver()
        s.parameters.max_time_in_seconds = float(time_limit)
        s.parameters.num_search_workers = int(workers)
        s.parameters.log_search_progress = bool(log)
        return s, s.Solve(model)

    solver = None
    st = None
    stage_values = []

    for si, (name, expr) in enumerate(stages):
        # 희소 셀 우선 단계 직전에 '셀별 최소 달성률' 하한을 건다
        guard = False
        if name == "scarcity" and min_fill and min_fill > 0:
            for k, tgt in main_map.items():
                keep = int(math.ceil(tgt * float(min_fill)))
                cap_short = max(0, tgt - keep)
                if cap_short < tgt:
                    model.Add(short_main[k] <= cap_short)
                    guard = True

        model.Minimize(expr)
        s_new, st_new = _run()

        if st_new not in (cp_model.OPTIMAL, cp_model.FEASIBLE):
            if si == 0:
                # main_hard 또는 편차 밴드로 해가 없어졌을 수 있다 → 완화 재시도
                banded = (ex_overflow and not ex_tol_unlimited) or (
                    ex_as_target and not ex_overflow)
                tag = ("BAND_INFEASIBLE" if banded
                       else ("HARD_INFEASIBLE" if main_hard else None))
                return (tag, st_new, s_new, n_vars, short_main, short_ex,
                        over_ex, stage_values)
            if guard:
                # 하한 때문에 해가 사라졌다 → 호출부가 하한 없이 재시도
                return ("GUARD_INFEASIBLE", st_new, solver, n_vars,
                        short_main, short_ex, over_ex, stage_values)
            break   # 이전 단계 해를 그대로 사용

        solver, st = s_new, st_new
        if st != cp_model.OPTIMAL:
            # 최적성이 증명되지 않았으면 값을 고정하는 것 자체가 위험하다
            return (True, st, solver, n_vars, short_main, short_ex,
                    over_ex, stage_values)

        val = int(round(solver.ObjectiveValue()))
        stage_values.append((name, val))
        if si < len(stages) - 1:
            model.Add(expr == val)

    return True, st, solver, n_vars, short_main, short_ex, over_ex, stage_values


# ==============================================================================
# 2-B. 추가 수집 지시서 (부족분을 어떤 구성으로 보충해야 하는가)
# ==============================================================================
@dataclass
class RecruitPlan:
    """추가 수집 계획."""
    feasible: bool = False
    total_needed: int = 0                  # 총 추가 수집 인원
    by_cell: dict = field(default_factory=dict)     # 메인 셀 -> 필요 인원
    rows: list = field(default_factory=list)        # 상세 지시 [{cell, pattern, n}]
    ex_dev_after: list = field(default_factory=list)  # 계획 반영 후 추가 쿼터 편차
    max_dev_after: int = 0
    n_patterns: int = 0
    solve_sec: float = 0.0
    notes: list = field(default_factory=list)


def plan_recruitment(m_keys, ex_keys_list, main_map, ex_maps,
                     unlisted="free", ex_as_target=True,
                     ex_tol_abs=0, ex_tol_pct=0.0, ex_tol_unlimited=True,
                     time_limit=30.0, workers=8, max_patterns=200):
    """
    메인 쿼터를 100% 채우려면 **어떤 조건의 응답자를 몇 명 더 수집해야 하는지**
    역산한다. 실사팀에 넘기는 추가 수집 지시서로 쓸 수 있다.

    방법
    ----
    기존 응답자 풀에 '아직 존재하지 않는 응답자' 변수를 추가해서 다시 푼다.

        n_p            기존 프로파일 p 에서 뽑을 인원   (0 ≤ n_p ≤ 보유수)
        r_(c, pat)     메인 셀 c 에서 추가 수집할 인원   (상한 없음)

        메인 : Σ_{p∈c} n_p + Σ_pat r_(c,pat) = T_c        (부족 0 = 하드)
        추가 : Σ n_p + Σ r + 부족 − 초과 = C_jk

        1단계  min Σ r        추가 수집 인원 최소화 (기존 표본을 최대한 활용)
        2단계  min Σ 편차      남는 추가 쿼터 편차 최소화
        3단계  min max 편차율  편차를 고르게 분산

    pat(패턴)은 '추가 쿼터 키 조합'이다. 후보는 데이터에 실제로 관측된 조합만
    쓴다. 모집단에 존재하지 않는 조합을 수집하라고 지시하면 의미가 없기 때문이다.

    반환: RecruitPlan
    """
    t0 = time.perf_counter()
    notes = []

    profiles = build_profiles(m_keys, ex_keys_list, main_map, ex_maps,
                              ex_as_target=True, unlisted=unlisted)
    if not profiles:
        return RecruitPlan(feasible=False, notes=["선택 가능한 응답자가 없습니다."])

    # --- 후보 패턴 : 데이터에 관측된 추가 키 조합 (빈도순) ---
    pat_count = collections.Counter()
    for (_mk, sig_ex), rows in profiles:
        pat_count[sig_ex] += len(rows)
    # 목표 0(금지) 키를 포함한 패턴은 수집 대상이 될 수 없다
    def _ok(sig_ex):
        for j, keys in enumerate(sig_ex):
            e_map = ex_maps[j] if j < len(ex_maps) else None
            if not e_map:
                continue
            for k in keys:
                if e_map.get(k, 0) <= 0:
                    return False
        return True
    patterns = [p for p, _c in pat_count.most_common() if _ok(p)][:max_patterns]
    if not patterns:
        patterns = [tuple(() for _ in ex_maps)]

    model = cp_model.CpModel()
    n_vars = [model.NewIntVar(0, len(rows), f"n{i}")
              for i, (_sig, rows) in enumerate(profiles)]

    by_main = collections.defaultdict(list)
    for i, ((mk, _), _rows) in enumerate(profiles):
        by_main[mk].append(i)

    # 추가 수집 변수 : (메인 셀, 패턴)
    r_vars = {}
    for k, tgt in main_map.items():
        if tgt <= 0:
            continue
        for pi, pat in enumerate(patterns):
            r_vars[(k, pi)] = model.NewIntVar(0, tgt, f"r{len(r_vars)}")

    # 메인 쿼터 : 기존 + 추가수집 = 목표 (부족 0)
    for k, tgt in main_map.items():
        if tgt <= 0:
            continue
        model.Add(sum(n_vars[i] for i in by_main.get(k, []))
                  + sum(r_vars[(k, pi)] for pi in range(len(patterns))) == tgt)

    # 추가 쿼터 : 기존 + 추가수집 + 부족 − 초과 = 목표
    by_ex = collections.defaultdict(list)
    for i, ((_mk, sig_ex), _rows) in enumerate(profiles):
        for j, keys in enumerate(sig_ex):
            for k in keys:
                by_ex[(j, k)].append(i)

    short_ex, over_ex = {}, {}
    total_target = sum(v for v in main_map.values() if v > 0)
    for j, e_map in enumerate(ex_maps):
        if not e_map:
            continue
        for k, cap in e_map.items():
            if cap <= 0:
                continue
            terms = [n_vars[i] for i in by_ex.get((j, k), [])]
            for (ck, pi), rv in r_vars.items():
                if k in patterns[pi][j]:
                    terms.append(rv)
            if ex_tol_unlimited:
                bs, bo = cap, total_target
            else:
                band = max(int(ex_tol_abs), int(math.ceil(cap * float(ex_tol_pct))))
                bs, bo = min(cap, band), band
            sv = model.NewIntVar(0, bs, f"sE{len(short_ex)}")
            ov = model.NewIntVar(0, bo, f"oE{len(over_ex)}")
            short_ex[(j, k)] = sv
            over_ex[(j, k)] = ov
            model.Add(sum(terms) + sv - ov == cap)

    def _run(tl):
        sv = cp_model.CpSolver()
        sv.parameters.max_time_in_seconds = float(tl)
        sv.parameters.num_search_workers = int(workers)
        return sv, sv.Solve(model)

    # 1단계 : 추가 수집 인원 최소화
    obj1 = sum(r_vars.values())
    model.Minimize(obj1)
    solver, st = _run(time_limit)
    if st not in (cp_model.OPTIMAL, cp_model.FEASIBLE):
        return RecruitPlan(
            feasible=False, n_patterns=len(patterns),
            solve_sec=time.perf_counter() - t0,
            notes=["허용 편차 안에서는 추가 수집으로도 목표를 맞출 수 없습니다. "
                   "편차 한계를 넓히거나 목표를 조정해야 합니다."])
    need = int(round(solver.ObjectiveValue()))
    if st == cp_model.OPTIMAL:
        model.Add(obj1 == need)
        # 2단계 : 남는 편차 최소화
        dev = sum(short_ex.values()) + sum(over_ex.values())
        model.Minimize(dev)
        s2, st2 = _run(time_limit)
        if st2 in (cp_model.OPTIMAL, cp_model.FEASIBLE):
            solver, st = s2, st2
        if st2 == cp_model.OPTIMAL:
            model.Add(dev == int(round(solver.ObjectiveValue())))
            # 3단계 : 남는 편차를 여러 항목에 고르게 분산 (제곱 편차 최소화)
            pairs = []
            for (j, k), sv in short_ex.items():
                cap = ex_maps[j].get(k, 1)
                pairs.append(([sv, over_ex[(j, k)]], cap, cap + total_target))
            sq_terms = _squared_dev_terms(model, pairs, tag="r")
            if sq_terms:
                model.Minimize(sum(sq_terms))
                s3, st3 = _run(min(20.0, time_limit))
                if st3 in (cp_model.OPTIMAL, cp_model.FEASIBLE):
                    solver, st = s3, st3
    else:
        notes.append("시간 제한 내에 최적성을 증명하지 못했습니다. "
                     "지시서는 유효하지만 더 적은 인원으로 가능할 수 있습니다.")

    # --- 결과 정리 ---
    plan = RecruitPlan(feasible=True, n_patterns=len(patterns), notes=notes)
    for (k, pi), rv in r_vars.items():
        v = solver.Value(rv)
        if v <= 0:
            continue
        plan.by_cell[k] = plan.by_cell.get(k, 0) + v
        plan.rows.append({
            'cell': k,
            'pattern': {j: patterns[pi][j] for j in range(len(ex_maps))
                        if patterns[pi][j]},
            'n': v,
        })
    plan.total_needed = sum(plan.by_cell.values())
    plan.rows.sort(key=lambda d: (-d['n'], repr(d['cell'])))

    plan.ex_dev_after = [{} for _ in ex_maps]
    mx = 0
    for (j, k), sv in short_ex.items():
        d = -solver.Value(sv) + solver.Value(over_ex[(j, k)])
        if d:
            plan.ex_dev_after[j][k] = d
            mx = max(mx, abs(d))
    plan.max_dev_after = mx
    plan.solve_sec = time.perf_counter() - t0
    return plan


# ==============================================================================
# 3. 공개 API
# ==============================================================================
def solve_quota_ilp(m_keys, ex_keys_list, main_map, ex_maps, indices,
                    weights=None, priority=True, balance=False, min_fill=0.0,
                    time_limit=30.0, workers=8,
                    diagnose=True, max_value_probes=8, rng=None, tiebreak=None,
                    ex_as_target=False, ex_weights=None, unlisted="free",
                    balance_relative=True, main_hard=False,
                    ex_overflow=False, overflow_weight=1,
                    ex_tol_abs=0, ex_tol_pct=0.0, ex_tol_unlimited=False):
    """
    쿼터 할당 최적화.

      m_keys       : 행별 메인 키 (튜플)
      ex_keys_list : 그룹별 [행별 키 리스트]
      main_map     : {메인 키: 목표}
      ex_maps      : [{추가 키: 목표 또는 상한}, ...]
      indices      : 행별 인덱스 라벨 (df.index.to_numpy())

      ex_as_target : True 면 추가 쿼터를 **목표**로 취급해 부족도 최소화한다
                     (초과는 두 경우 모두 금지)
      ex_weights   : {그룹 번호: 가중치}. 추가 쿼터끼리 우선순위를 줄 때 사용.
      unlisted     : 목표 목록에 없는 값의 처리. 'free'(제약 없음) / 'forbid'(제외)
      priority     : 통과 인원을 고정한 뒤 희소한 메인 셀을 우선 채운다
      balance      : 남은 부족분을 고르게 분산한다 (priority 보다 뒤 단계)
      min_fill     : 0~1. 어떤 메인 셀도 목표의 이 비율 미만이 되지 않게 한다
      main_hard    : True 면 메인 쿼터를 **하드 쿼터**로 걸어 셀마다 정확히 목표
                     인원을 뽑는다. 물리적으로 불가능하면 해가 사라지므로,
                     그때는 자동으로 소프트로 되돌리고 notes 에 사유를 남긴다.
      ex_overflow  : True 면 추가 쿼터가 목표에서 벗어나는 것을 허용한다.
                     총 선정 인원은 메인 쿼터가 정하므로 그대로이고, 개별 항목만
                     목표 위아래로 흔들린다.
                       50/50/50/50 (합 200)  →  55/45/47/53 (합 200)
                     False 면 초과가 금지되어 편차를 흡수할 수 없고, 그 대신
                     메인 쿼터가 미달하게 된다.
      ex_tol_abs   : 항목별 허용 편차(명). 예: 5 → 각 항목 ±5명까지
      ex_tol_pct   : 항목별 허용 편차(비율). 예: 0.1 → 각 항목 ±10%까지
                     (두 값 중 큰 쪽이 적용된다. 둘 다 0 이면 편차 0 = 하드)
      ex_tol_unlimited : True 면 편차 한계 없이 최소화만 한다
      overflow_weight : 초과 1명을 부족 몇 명만큼 싫어할지. 기본 1 (동등).
                     주의: 단일응답 그룹에서는 다음 항등식 때문에 이 값이 결과를
                     바꾸지 못한다.
                         Σ초과 − Σ부족 = 선정인원 − Σ목표   (우변 고정)
                     따라서 Σ초과 = 고정값 + Σ부족 이 되어, 어떤 가중치를 줘도
                     '부족 0 + 초과 고정값'이 항상 최적이다. 이 값이 실제로
                     의미를 갖는 것은 응답자당 키 개수가 다른 복수응답 그룹이다.
      balance_relative : 부족 분산 기준. True 면 '목표 대비 부족률', False 면 인원수
      tiebreak     : 행별 실수 배열. 프로파일 내부에서 값이 큰 쪽을 먼저 선택

    ex_weights 에 대한 주의
    ------------------------
    추가 쿼터 그룹이 단일응답(응답자당 키 1개)이면 다음 항등식이 성립한다.

        그룹 부족 합계 = 메인 부족 + (Σ 추가목표 − 메인목표 합계)

    즉 그룹의 **부족 총량은 가중치와 무관하게 산술적으로 고정**되며,
    ex_weights 로 그룹 간에 부족을 옮길 수 없다. 이때 실제로 제어할 수 있는
    것은 '어느 항목이 그 부족을 감당할지'이고, 그것은 balance 단계가 결정한다.
    ex_weights 가 의미를 갖는 경우는 복수응답 그룹(응답자당 키 개수가 다름)이다.

    같은 항등식에서 따라오는 결론 : 선택 인원을 늘리면 그룹 부족도 같이 줄어들기
    때문에, '메인 우선 → 추가' 순서의 사전식 최적화는 추가 쿼터를 희생시키지
    않는다. 두 목표가 충돌하지 않으므로 순서를 고민할 필요가 없다.

    반환: IlpSolution
    """
    t0 = time.perf_counter()
    target_total = sum(main_map.values())

    pre = preflight_targets(m_keys, ex_keys_list, main_map, ex_maps,
                            ex_as_target=ex_as_target, unlisted=unlisted,
                            ex_overflow=ex_overflow)

    profiles = build_profiles(m_keys, ex_keys_list, main_map, ex_maps,
                              ex_as_target=ex_as_target, unlisted=unlisted)
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
                arithmetic=pre,
            ),
            n_profiles=0, n_rows_considered=0,
            solve_sec=time.perf_counter() - t0,
            ex_as_target=ex_as_target,
            main_short_total=target_total,
            ex_short_total=[sum(m.values()) if m else 0 for m in ex_maps],
            ex_over_total=[0 for _ in ex_maps],
        )

    notes = []
    kw = dict(weights=weights, ex_as_target=ex_as_target, ex_weights=ex_weights,
              overflow_weight=overflow_weight,
              priority=priority, balance=balance, balance_relative=balance_relative,
              time_limit=time_limit, workers=workers)

    def _try(**over):
        base = dict(ex_overflow=ex_overflow, ex_tol_abs=ex_tol_abs,
                    ex_tol_pct=ex_tol_pct, ex_tol_unlimited=ex_tol_unlimited,
                    main_hard=main_hard, min_fill=min_fill)
        base.update(over)
        return _solve_core(profiles, main_map, ex_maps, **base, **kw), base

    # 완화 순서 : ① 추가 쿼터 편차 한계 → ② 메인 하드
    # 메인 쿼터가 총량을 정의하므로 메인을 마지막에 풀어준다.
    (res, used) = _try()
    ok = res[0]

    if ok == "BAND_INFEASIBLE":
        band_txt = ("±{}명".format(ex_tol_abs) if ex_tol_abs else
                    "±{:.0%}".format(ex_tol_pct) if ex_tol_pct else "편차 0(정확히 맞춤)")
        notes.append(
            f"추가 쿼터 허용 편차 {band_txt} 안에서는 해가 존재하지 않습니다. "
            "메인 쿼터를 채우려면 추가 쿼터가 그보다 더 벗어나야 합니다. "
            "한계를 풀고 편차를 최소화하는 방식으로 다시 계산했습니다.")
        ex_overflow, ex_tol_unlimited = True, True
        (res, used) = _try(ex_overflow=True, ex_tol_unlimited=True)
        ok = res[0]

    if ok == "HARD_INFEASIBLE":
        notes.append(
            "메인 쿼터를 하드 쿼터로 걸면 해가 존재하지 않습니다. 셀별 목표를 "
            "정확히 채우는 것이 물리적으로 불가능하다는 뜻입니다. 부족을 허용하는 "
            "방식으로 다시 계산했으니, 아래 부족 분석에서 어느 셀이 불가능한지 "
            "확인하세요.")
        main_hard = False
        (res, used) = _try(main_hard=False)
        ok = res[0]

    (ok, st, solver, n_vars, short_main, short_ex, over_ex, stage_values) = res

    if ok == "GUARD_INFEASIBLE":
        notes.append(
            f"최소 달성률 {min_fill:.0%} 를 모든 셀에서 만족시킬 수 없어 "
            "하한 없이 다시 계산했습니다. 일부 셀이 크게 미달할 수 있습니다.")
        (res, used) = _try(min_fill=0.0)
        (ok, st, solver, n_vars, short_main, short_ex, over_ex, stage_values) = res

    if not ok:
        return IlpSolution(
            status=solver.StatusName(st) if solver else "INFEASIBLE",
            proven_optimal=False, selected=[], total=0,
            target_total=target_total, main_actual={},
            ex_actual=[{} for _ in ex_maps],
            diagnosis=Diagnosis(arithmetic=pre),
            n_profiles=len(profiles), n_rows_considered=n_considered,
            solve_sec=time.perf_counter() - t0, ex_as_target=ex_as_target,
        )

    # ------------------------------------------------------------------
    # 해를 개인 단위로 펼치기
    #   프로파일 구성원은 서로 완전히 교환 가능하므로 누구를 뽑을지는
    #   최적 인원수와 무관하다.
    #     tiebreak → 값이 큰 쪽부터 / rng → 무작위 / 없으면 데이터 순서
    # ------------------------------------------------------------------
    tb = None
    if tiebreak is not None:
        import numpy as _np
        tb = _np.asarray(tiebreak, dtype=float)

    selected = []
    for idx, (_sig, rows) in enumerate(profiles):
        take = solver.Value(n_vars[idx])
        if take <= 0:
            continue
        if take >= len(rows):
            pick = rows
        elif tb is not None:
            pick = sorted(rows, key=lambda i: -tb[i])[:take]
        elif rng is not None:
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

    main_short_total = target_total - total
    ex_short_total, ex_over_total = [], []
    for j, e_map in enumerate(ex_maps):
        if not e_map:
            ex_short_total.append(0)
            ex_over_total.append(0)
            continue
        ex_short_total.append(sum(max(0, c - ex_actual[j].get(k, 0))
                                  for k, c in e_map.items()))
        ex_over_total.append(sum(max(0, ex_actual[j].get(k, 0) - c)
                                 for k, c in e_map.items()))

    # ------------------------------------------------------------------
    # 진단
    # ------------------------------------------------------------------
    diag = Diagnosis(arithmetic=pre)
    diag.ex_short = [{} for _ in ex_maps]
    diag.ex_over = [{} for _ in ex_maps]
    diag.ex_avail = [{} for _ in ex_maps]
    diag.ex_reason = [{} for _ in ex_maps]

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
                p = tgt - avail.get(k, 0)
                diag.main_reason[k] = (
                    f"⚠️ 물리적 부족 (보유 {avail.get(k, 0)}명)" if short <= p
                    else f"⚠️+⚔️ 물리적 {p}명 + 경합 {short - p}명")
            else:
                diag.main_reason[k] = "⚔️ 경합 부족 (추가 쿼터에 막힘)"

        # 추가 쿼터 부족 사유
        group_sum_bad = {d['group'] for d in pre
                         if d['kind'] in ('group_sum_low', 'group_sum_high')}
        for j, e_map in enumerate(ex_maps):
            if not e_map:
                continue
            avail_k = collections.Counter()
            for (_mk, sig_ex), rows in profiles:
                for k in sig_ex[j]:
                    avail_k[k] += len(rows)
            for k, cap in e_map.items():
                used = ex_actual[j].get(k, 0)
                diag.ex_avail[j][k] = avail_k.get(k, 0)
                if cap > 0 and used >= cap:
                    diag.binding.append({'group': j, 'key': k,
                                         'cap': cap, 'used': used})
                short = cap - used
                if short < 0:
                    diag.ex_over[j][k] = -short
                if short <= 0:
                    continue
                diag.ex_short[j][k] = short
                if avail_k.get(k, 0) < cap:
                    diag.ex_reason[j][k] = f"⚠️ 물리적 부족 (보유 {avail_k.get(k, 0)}명)"
                elif j in group_sum_bad:
                    diag.ex_reason[j][k] = "⚖️ 구조적 (그룹 목표 합계 불일치)"
                else:
                    diag.ex_reason[j][k] = "⚔️ 경합 부족 (다른 쿼터와 충돌)"

        # 미달 원인 정량화
        if total < target_total:
            active = [j for j, m in enumerate(ex_maps) if m]

            # (a) 그룹 단위 완화 : 이 그룹의 초과 금지를 풀면 몇 명 더 뽑히는가
            for j in active:
                r = _solve_core(profiles, main_map, ex_maps, weights=weights,
                                ex_as_target=False, priority=False, balance=False,
                                skip_groups=frozenset([j]),
                                time_limit=min(10.0, time_limit), workers=workers)
                if r[0]:
                    gain = sum(r[2].Value(v) for v in r[3]) - total
                    if gain > 0:
                        diag.group_relax_gain[j] = gain

            # (b) 값 단위 민감도 : 병목 한도를 1 늘리면 몇 명 더 뽑히는가
            probes = sorted(diag.binding, key=lambda d: -d['cap'])[:max_value_probes]
            for b in probes:
                r = _solve_core(profiles, main_map, ex_maps, weights=weights,
                                ex_as_target=False, priority=False, balance=False,
                                cap_bonus={(b['group'], b['key']): 1},
                                time_limit=min(5.0, time_limit), workers=workers)
                if r[0]:
                    gain = sum(r[2].Value(v) for v in r[3]) - total
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
        notes=notes,
        ex_as_target=ex_as_target,
        main_short_total=main_short_total,
        ex_short_total=ex_short_total,
        ex_over_total=ex_over_total,
        stage_values=stage_values,
        all_satisfied=(main_short_total == 0 and sum(ex_short_total) == 0
                       and sum(ex_over_total) == 0),
        main_hard=main_hard,
        ex_overflow=ex_overflow,
    )

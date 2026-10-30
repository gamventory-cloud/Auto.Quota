"""쿼터 솔루션 단계별 풀이 — 뒤 단계를 끝내지 못했을 때 failed_stages·최적 증명 표시 검증.

    python -m unittest tests.test_quota_ilp_stages      (저장소 맨 위 폴더에서)

_solve_core 는 main → ex → scarcity → balance → spread 순서로 한 단계씩 풀고 값을 고정한다.
뒤 단계가 UNKNOWN(시간 안에 해 없음) 이면 앞 단계 해를 그대로 쓰고, FEASIBLE(최적 미증명) 이면
그 해로 일찍 끝낸다. 예전에는 앞 단계 상태 OPTIMAL 을 돌려줘 proven_optimal=True 로 보였다
(dp-skills#204 7-5, Auto.Quota#8 검수 반영 2026-10-30). CpSolver.Solve 를 바꿔 n번째 단계 상태를 강제한다.
"""
import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

try:
    import numpy as np
    from ortools.sat.python import cp_model
    import quota_ilp
except ImportError as e:                       # ortools 가 없는 PC 에서는 건너뜀
    quota_ilp = None
    _SKIP = f"ortools/numpy 없음: {e}"

STAGES = ["main", "ex", "scarcity", "balance", "spread"]


def _data():
    # 성별x연령 4셀, 한 셀은 표본 부족 → scarcity·balance·spread 단계가 모두 생긴다
    cells = {("1", "20"): 8, ("1", "30"): 3, ("2", "20"): 6, ("2", "30"): 9}
    m_keys, ex_keys = [], []
    for k, n in cells.items():
        for i in range(n):
            m_keys.append(k)
            ex_keys.append(("A" if i % 2 else "B",))
    tgt = {k: 5 for k in cells}
    return m_keys, [ex_keys], tgt, [{("A",): 9, ("B",): 9}], np.arange(len(m_keys))


@unittest.skipIf(quota_ilp is None, _SKIP if quota_ilp is None else "")
class SolveStageTest(unittest.TestCase):
    def solve(self, fail_at=0, status=None):
        real = cp_model.CpSolver
        calls = {"n": 0}

        class Forced(real):
            def Solve(self, model, *a, **kw):
                calls["n"] += 1
                if calls["n"] != fail_at:
                    return real.Solve(self, model, *a, **kw)
                if status == cp_model.UNKNOWN:
                    return cp_model.UNKNOWN                  # 해를 하나도 못 찾음
                st = real.Solve(self, model, *a, **kw)       # 해는 있지만 최적 증명 못 함
                assert st == cp_model.OPTIMAL
                return cp_model.FEASIBLE

        quota_ilp.cp_model.CpSolver = Forced
        try:
            m_keys, ex_keys, tgt, ex_maps, idx = _data()
            return quota_ilp.solve_quota_ilp(m_keys, ex_keys, tgt, ex_maps, idx,
                                             priority=True, balance=True, time_limit=5,
                                             workers=1, diagnose=False, ex_as_target=True,
                                             ex_overflow=True, ex_tol_unlimited=True)
        finally:
            quota_ilp.cp_model.CpSolver = real

    def test_all_stages_optimal(self):
        sol = self.solve()
        self.assertEqual([n for n, _ in sol.stage_values], STAGES)
        self.assertTrue(sol.proven_optimal)
        self.assertEqual(sol.failed_stages, [])
        self.assertFalse(any("최적으로 정해지지" in n for n in sol.notes))

    def test_later_stage_unknown(self):
        base = self.solve()
        for k in range(2, len(STAGES) + 1):
            with self.subTest(stage=STAGES[k - 1]):
                sol = self.solve(k, cp_model.UNKNOWN)
                self.assertEqual([n for n, _ in sol.stage_values], STAGES[:k - 1])
                self.assertEqual(sol.failed_stages,
                                 [(STAGES[k - 1], "UNKNOWN")] + [(n, "NOT_RUN") for n in STAGES[k:]])
                self.assertFalse(sol.proven_optimal)
                self.assertEqual(sol.status, "OPTIMAL")              # 앞 단계 해 — 상태는 그대로
                self.assertEqual(sol.total, base.total)
                self.assertTrue(any("풀지 못해 앞 단계" in n for n in sol.notes), sol.notes)

    def test_later_stage_feasible(self):
        base = self.solve()
        for k in range(2, len(STAGES) + 1):
            with self.subTest(stage=STAGES[k - 1]):
                sol = self.solve(k, cp_model.FEASIBLE)
                self.assertEqual([n for n, _ in sol.stage_values], STAGES[:k - 1])
                self.assertEqual(sol.failed_stages,
                                 [(STAGES[k - 1], "FEASIBLE")] + [(n, "NOT_RUN") for n in STAGES[k:]])
                self.assertFalse(sol.proven_optimal)
                self.assertEqual(sol.status, "FEASIBLE")
                self.assertEqual(sol.total, base.total)
                self.assertTrue(any("최적으로 확정하지 못해" in n for n in sol.notes), sol.notes)


if __name__ == "__main__":
    unittest.main()

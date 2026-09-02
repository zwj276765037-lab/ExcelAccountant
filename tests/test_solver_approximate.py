from __future__ import annotations

import itertools
import threading

from excel_accountant.models import ScaledProblem, SolveStatus
from excel_accountant.solver_approximate import solve_approximate


def _metrics(problem: ScaledProblem, allocation: tuple[int, ...]) -> tuple[int, int, int, int]:
    actuals = [0] * len(problem.targets)
    used = 0
    for cell_index, target_index in enumerate(allocation):
        if target_index >= 0:
            actuals[target_index] += problem.amounts[cell_index]
            used += 1
    differences = [actual - target for actual, target in zip(actuals, problem.targets)]
    exact_count = sum(value == 0 for value in differences)
    return (
        -exact_count,
        sum(abs(value) for value in differences),
        max((abs(value) for value in differences), default=0),
        used,
    )


def _brute_best_metrics(problem: ScaledProblem) -> tuple[int, int, int, int]:
    allocations = itertools.product(
        range(-1, len(problem.targets)),
        repeat=len(problem.amounts),
    )
    return min(_metrics(problem, allocation) for allocation in allocations)


def _assert_disjoint(solution) -> None:
    used: set[int] = set()
    for assignment in solution.assignments:
        assert not used.intersection(assignment.cell_indices)
        used.update(assignment.cell_indices)


def test_nearest_single_target_is_reported() -> None:
    problem = ScaledProblem(1, 1, (4, 6), (5,))
    outcome = solve_approximate(problem, max_solutions=2, time_limit_seconds=5)
    assert outcome.status == SolveStatus.APPROXIMATE_READY
    assert len(outcome.approximate_solutions) == 2
    assert all(item.total_absolute_difference == 1 for item in outcome.approximate_solutions)


def test_multiple_targets_remain_disjoint() -> None:
    problem = ScaledProblem(1, 1, (5, 7, 8), (6, 9))
    outcome = solve_approximate(problem, max_solutions=5, time_limit_seconds=5)
    assert outcome.approximate_solutions
    for solution in outcome.approximate_solutions:
        _assert_disjoint(solution)


def test_exact_target_count_has_first_priority() -> None:
    problem = ScaledProblem(1, 1, (5, 100), (5, 6))
    outcome = solve_approximate(problem, max_solutions=1, time_limit_seconds=5)
    solution = outcome.approximate_solutions[0]
    assert solution.exact_target_count == 1
    assert any(item.difference == 0 for item in solution.assignments)


def test_first_solution_matches_brute_force_lexicographic_optimum() -> None:
    problem = ScaledProblem(1, 1, (2, 4, 7, -1), (5, 8))
    expected = _brute_best_metrics(problem)
    outcome = solve_approximate(problem, max_solutions=1, time_limit_seconds=5)
    solution = outcome.approximate_solutions[0]
    actual = (
        -solution.exact_target_count,
        solution.total_absolute_difference,
        solution.maximum_absolute_difference,
        sum(len(item.cell_indices) for item in solution.assignments),
    )
    assert actual == expected


def test_alternative_solutions_are_unique_and_ordered() -> None:
    problem = ScaledProblem(1, 1, (4, 6, 7), (5,))
    outcome = solve_approximate(problem, max_solutions=5, time_limit_seconds=5)
    signatures = {
        tuple(item.cell_indices for item in solution.assignments)
        for solution in outcome.approximate_solutions
    }
    assert len(signatures) == len(outcome.approximate_solutions)
    metrics = [
        (
            -solution.exact_target_count,
            solution.total_absolute_difference,
            solution.maximum_absolute_difference,
            sum(len(item.cell_indices) for item in solution.assignments),
        )
        for solution in outcome.approximate_solutions
    ]
    assert metrics == sorted(metrics)


def test_pre_cancelled_approximate_search_stops() -> None:
    event = threading.Event()
    event.set()
    problem = ScaledProblem(1, 1, (1, 2, 3), (10,))
    outcome = solve_approximate(problem, cancel_event=event)
    assert outcome.status == SolveStatus.CANCELLED
    assert not outcome.approximate_solutions

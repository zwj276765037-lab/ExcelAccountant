from __future__ import annotations

import itertools
import random
import threading

from excel_accountant.models import ScaledProblem, SolveStatus
from excel_accountant.solver_exact import solve_exact


def _assert_valid(problem: ScaledProblem, outcome) -> None:
    for solution in outcome.exact_solutions:
        used: set[int] = set()
        assert len(solution.assignments) == len(problem.targets)
        for assignment in solution.assignments:
            assert assignment.cell_indices
            assert not used.intersection(assignment.cell_indices)
            used.update(assignment.cell_indices)
            assert (
                sum(problem.amounts[index] for index in assignment.cell_indices)
                == problem.targets[assignment.target_index]
            )


def _canonical(assignments: tuple[tuple[int, ...], ...], targets: tuple[int, ...]):
    grouped: dict[int, list[tuple[int, ...]]] = {}
    for index, cells in enumerate(assignments):
        grouped.setdefault(targets[index], []).append(tuple(sorted(cells)))
    return tuple(
        (target, tuple(sorted(groups))) for target, groups in sorted(grouped.items())
    )


def _brute_force(problem: ScaledProblem) -> set[tuple]:
    results: set[tuple] = set()
    target_count = len(problem.targets)
    for allocation in itertools.product(range(-1, target_count), repeat=len(problem.amounts)):
        assignments = tuple(
            tuple(index for index, target_index in enumerate(allocation) if target_index == target)
            for target in range(target_count)
        )
        if any(not cells for cells in assignments):
            continue
        if all(
            sum(problem.amounts[index] for index in assignments[target])
            == problem.targets[target]
            for target in range(target_count)
        ):
            results.add(_canonical(assignments, problem.targets))
    return results


def test_multiple_targets_are_disjoint() -> None:
    problem = ScaledProblem(1, 1, (100, 200, 300, 400, 500), (300, 500))
    outcome = solve_exact(problem, max_solutions=20, time_limit_seconds=5)
    assert outcome.exact_solutions
    assert outcome.status == SolveStatus.EXACT_COMPLETE
    assert outcome.search_complete is True
    _assert_valid(problem, outcome)


def test_no_solution_is_proved() -> None:
    problem = ScaledProblem(1, 1, (100,), (50,))
    outcome = solve_exact(problem, time_limit_seconds=5)
    assert outcome.status == SolveStatus.NO_EXACT_PROVED
    assert outcome.search_complete is True
    assert not outcome.exact_solutions


def test_duplicate_targets_do_not_emit_label_swaps() -> None:
    problem = ScaledProblem(1, 1, (1, 1), (1, 1))
    outcome = solve_exact(problem, max_solutions=10, time_limit_seconds=5)
    assert outcome.status == SolveStatus.EXACT_COMPLETE
    assert len(outcome.exact_solutions) == 1


def test_negative_values_and_nonempty_zero_target() -> None:
    negative_problem = ScaledProblem(1, 1, (5, -2, 3), (1,))
    negative_outcome = solve_exact(negative_problem, time_limit_seconds=5)
    _assert_valid(negative_problem, negative_outcome)
    assert negative_outcome.exact_solutions

    zero_problem = ScaledProblem(1, 1, (1, -1), (0,))
    zero_outcome = solve_exact(zero_problem, time_limit_seconds=5)
    _assert_valid(zero_problem, zero_outcome)
    assert zero_outcome.exact_solutions


def test_empty_assignment_cannot_satisfy_zero() -> None:
    problem = ScaledProblem(1, 1, (5,), (0,))
    outcome = solve_exact(problem, time_limit_seconds=5)
    assert outcome.status == SolveStatus.NO_EXACT_PROVED


def test_solution_limit_is_reported() -> None:
    problem = ScaledProblem(1, 1, (1, 1, 1, 1), (2,))
    outcome = solve_exact(problem, max_solutions=2, time_limit_seconds=5)
    assert outcome.status == SolveStatus.EXACT_TRUNCATED
    assert len(outcome.exact_solutions) == 2


def test_pre_cancelled_search_stops_without_claiming_no_solution() -> None:
    event = threading.Event()
    event.set()
    problem = ScaledProblem(1, 1, (1, 2, 3), (3,))
    outcome = solve_exact(problem, cancel_event=event)
    assert outcome.status == SolveStatus.CANCELLED
    assert outcome.search_complete is False


def test_random_small_cases_match_brute_force() -> None:
    generator = random.Random(20260902)
    for _ in range(12):
        amounts = tuple(generator.randint(-3, 6) or 1 for _ in range(5))
        targets = (generator.randint(-2, 8), generator.randint(-2, 8))
        problem = ScaledProblem(1, 1, amounts, targets)
        expected = _brute_force(problem)
        outcome = solve_exact(problem, max_solutions=100, time_limit_seconds=5)
        actual = {
            _canonical(
                tuple(item.cell_indices for item in solution.assignments),
                problem.targets,
            )
            for solution in outcome.exact_solutions
        }
        assert outcome.search_complete is True
        assert actual == expected

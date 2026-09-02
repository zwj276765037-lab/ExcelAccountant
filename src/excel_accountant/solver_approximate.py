from __future__ import annotations

import threading
import time
from collections import defaultdict
from dataclasses import dataclass

from ortools.sat.python import cp_model

from .models import (
    ApproximateAssignment,
    ApproximateSolution,
    ScaledProblem,
    SolveOutcome,
    SolveStatus,
)
from .solver_exact import SolverInputError


@dataclass(slots=True)
class _ApproximateModel:
    model: cp_model.CpModel
    assignments: list[list[cp_model.IntVar]]
    actuals: list[cp_model.IntVar]
    differences: list[cp_model.IntVar]
    absolute_differences: list[cp_model.IntVar]
    exact_flags: list[cp_model.IntVar]
    maximum_difference: cp_model.IntVar


def solve_approximate(
    problem: ScaledProblem,
    *,
    max_solutions: int = 5,
    time_limit_seconds: float = 30.0,
    cancel_event: threading.Event | None = None,
) -> SolveOutcome:
    if not 1 <= max_solutions <= 5:
        raise SolverInputError("近似方案数量必须在 1 至 5 之间")
    if time_limit_seconds <= 0:
        raise SolverInputError("搜索时限必须大于 0 秒")

    started = time.monotonic()
    exclusions: list[tuple[bool, ...]] = []
    signatures: set[tuple] = set()
    solutions: list[ApproximateSolution] = []

    while len(solutions) < max_solutions:
        if cancel_event is not None and cancel_event.is_set():
            return _outcome(SolveStatus.CANCELLED, solutions, started, "搜索已取消。")
        remaining = time_limit_seconds - (time.monotonic() - started)
        if remaining <= 0:
            break
        built = _build_model(problem, exclusions)
        solver = _solve_lexicographically(
            built,
            remaining,
            cancel_event,
        )
        if solver is None:
            break
        solution, raw_assignment = _extract_solution(problem, built, solver)
        exclusions.append(raw_assignment)
        signature = _canonical_signature(solution, problem.targets)
        if signature not in signatures:
            signatures.add(signature)
            solutions.append(solution)

    if cancel_event is not None and cancel_event.is_set():
        status = SolveStatus.CANCELLED
        message = "搜索已取消。"
    elif solutions:
        status = SolveStatus.APPROXIMATE_READY
        message = f"已生成 {len(solutions)} 套仅供参考的近似方案。"
    else:
        status = SolveStatus.EXACT_NOT_FOUND_TIMEOUT
        message = "未能在当前时限内生成近似方案。"
    return _outcome(status, solutions, started, message)


def _build_model(
    problem: ScaledProblem,
    exclusions: list[tuple[bool, ...]],
) -> _ApproximateModel:
    model = cp_model.CpModel()
    target_count = len(problem.targets)
    assignments = [
        [
            model.new_bool_var(f"cell_{cell_index}_target_{target_index}")
            for target_index in range(target_count)
        ]
        for cell_index in range(len(problem.amounts))
    ]
    for cell_vars in assignments:
        model.add(sum(cell_vars) <= 1)

    minimum_actual = sum(value for value in problem.amounts if value < 0)
    maximum_actual = sum(value for value in problem.amounts if value > 0)
    actuals: list[cp_model.IntVar] = []
    differences: list[cp_model.IntVar] = []
    absolute_differences: list[cp_model.IntVar] = []
    absolute_bounds: list[int] = []
    exact_flags: list[cp_model.IntVar] = []

    for target_index, target in enumerate(problem.targets):
        actual = model.new_int_var(
            minimum_actual,
            maximum_actual,
            f"actual_{target_index}",
        )
        model.add(
            actual
            == sum(
                problem.amounts[cell_index] * assignments[cell_index][target_index]
                for cell_index in range(len(problem.amounts))
            )
        )
        minimum_difference = minimum_actual - target
        maximum_difference = maximum_actual - target
        difference = model.new_int_var(
            minimum_difference,
            maximum_difference,
            f"difference_{target_index}",
        )
        model.add(difference == actual - target)
        absolute_bound = max(abs(minimum_difference), abs(maximum_difference))
        absolute = model.new_int_var(
            0,
            absolute_bound,
            f"absolute_difference_{target_index}",
        )
        model.add_abs_equality(absolute, difference)
        exact = model.new_bool_var(f"exact_{target_index}")
        model.add(absolute == 0).only_enforce_if(exact)
        model.add(absolute >= 1).only_enforce_if(exact.negated())
        actuals.append(actual)
        differences.append(difference)
        absolute_differences.append(absolute)
        absolute_bounds.append(absolute_bound)
        exact_flags.append(exact)

    maximum_bound = max(absolute_bounds, default=0)
    maximum_difference_var = model.new_int_var(
        0,
        maximum_bound,
        "maximum_absolute_difference",
    )
    model.add_max_equality(maximum_difference_var, absolute_differences)

    flat_variables = [variable for row in assignments for variable in row]
    for excluded in exclusions:
        if len(excluded) != len(flat_variables):
            continue
        matching = [
            variable if value else variable.negated()
            for variable, value in zip(flat_variables, excluded, strict=True)
        ]
        model.add_bool_or([literal.negated() for literal in matching])

    return _ApproximateModel(
        model=model,
        assignments=assignments,
        actuals=actuals,
        differences=differences,
        absolute_differences=absolute_differences,
        exact_flags=exact_flags,
        maximum_difference=maximum_difference_var,
    )


def _solve_lexicographically(
    built: _ApproximateModel,
    time_limit_seconds: float,
    cancel_event: threading.Event | None,
) -> cp_model.CpSolver | None:
    started = time.monotonic()
    exact_count = sum(built.exact_flags)
    total_difference = sum(built.absolute_differences)
    used_count = sum(variable for row in built.assignments for variable in row)
    stages = (
        ("max", exact_count),
        ("min", total_difference),
        ("min", built.maximum_difference),
        ("min", used_count),
    )
    last_solver: cp_model.CpSolver | None = None

    for direction, expression in stages:
        remaining = time_limit_seconds - (time.monotonic() - started)
        if remaining <= 0:
            break
        if direction == "max":
            built.model.maximize(expression)
        else:
            built.model.minimize(expression)
        solver = cp_model.CpSolver()
        solver.parameters.max_time_in_seconds = remaining
        solver.parameters.num_search_workers = 8
        status = _solve_with_cancellation(
            solver,
            built.model,
            cancel_event,
        )
        if status not in (cp_model.OPTIMAL, cp_model.FEASIBLE):
            break
        last_solver = solver
        best_value = solver.value(expression)
        built.model.add(expression == best_value)
        built.model.clear_objective()
        if cancel_event is not None and cancel_event.is_set():
            break
    return last_solver


def _solve_with_cancellation(
    solver: cp_model.CpSolver,
    model: cp_model.CpModel,
    cancel_event: threading.Event | None,
) -> cp_model.CpSolverStatus:
    if cancel_event is None:
        return solver.solve(model)
    monitor_done = threading.Event()

    def monitor() -> None:
        while not monitor_done.wait(0.05):
            if cancel_event.is_set():
                solver.stop_search()
                return

    thread = threading.Thread(target=monitor, daemon=True)
    thread.start()
    try:
        return solver.solve(model)
    finally:
        monitor_done.set()
        thread.join(timeout=0.2)


def _extract_solution(
    problem: ScaledProblem,
    built: _ApproximateModel,
    solver: cp_model.CpSolver,
) -> tuple[ApproximateSolution, tuple[bool, ...]]:
    assignments: list[ApproximateAssignment] = []
    for target_index in range(len(problem.targets)):
        cell_indices = tuple(
            cell_index
            for cell_index, row in enumerate(built.assignments)
            if solver.value(row[target_index])
        )
        actual = solver.value(built.actuals[target_index])
        difference = solver.value(built.differences[target_index])
        assignments.append(
            ApproximateAssignment(
                target_index=target_index,
                cell_indices=cell_indices,
                actual=actual,
                difference=difference,
            )
        )
    absolute_values = [abs(item.difference) for item in assignments]
    solution = ApproximateSolution(
        assignments=tuple(assignments),
        exact_target_count=sum(value == 0 for value in absolute_values),
        total_absolute_difference=sum(absolute_values),
        maximum_absolute_difference=max(absolute_values, default=0),
    )
    raw = tuple(
        bool(solver.value(variable))
        for row in built.assignments
        for variable in row
    )
    return solution, raw


def _canonical_signature(
    solution: ApproximateSolution,
    targets: tuple[int, ...],
) -> tuple:
    grouped: dict[int, list[tuple]] = defaultdict(list)
    for assignment in solution.assignments:
        grouped[targets[assignment.target_index]].append(
            (
                tuple(sorted(assignment.cell_indices)),
                assignment.actual,
                assignment.difference,
            )
        )
    return tuple(
        (target, tuple(sorted(items)))
        for target, items in sorted(grouped.items())
    )


def _outcome(
    status: SolveStatus,
    solutions: list[ApproximateSolution],
    started: float,
    message: str,
) -> SolveOutcome:
    return SolveOutcome(
        status=status,
        approximate_solutions=tuple(solutions),
        search_complete=False,
        wall_time_seconds=time.monotonic() - started,
        message=message,
    )

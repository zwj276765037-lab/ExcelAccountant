from __future__ import annotations

import threading
import time
from collections import defaultdict

from ortools.sat.python import cp_model

from .models import (
    ExactSolution,
    ScaledProblem,
    SolveOutcome,
    SolveStatus,
    TargetAssignment,
)


class SolverInputError(ValueError):
    """Raised when exact solver limits are invalid."""


def solve_exact(
    problem: ScaledProblem,
    *,
    max_solutions: int = 20,
    time_limit_seconds: float = 60.0,
    cancel_event: threading.Event | None = None,
) -> SolveOutcome:
    if not 1 <= max_solutions <= 100:
        raise SolverInputError("精确方案数量必须在 1 至 100 之间")
    if time_limit_seconds <= 0:
        raise SolverInputError("搜索时限必须大于 0 秒")

    started = time.monotonic()
    model = cp_model.CpModel()
    variables = [
        [
            model.new_bool_var(f"cell_{cell_index}_target_{target_index}")
            for target_index in range(len(problem.targets))
        ]
        for cell_index in range(len(problem.amounts))
    ]

    for cell_vars in variables:
        model.add(sum(cell_vars) <= 1)
    for target_index, target in enumerate(problem.targets):
        assigned = [variables[index][target_index] for index in range(len(variables))]
        model.add(
            sum(
                problem.amounts[index] * variables[index][target_index]
                for index in range(len(variables))
            )
            == target
        )
        model.add(sum(assigned) >= 1)

    unique_solutions: list[ExactSolution] = []
    signatures: set[tuple] = set()
    search_complete = False
    terminal_status = SolveStatus.EXACT_NOT_FOUND_TIMEOUT

    while len(unique_solutions) < max_solutions:
        if cancel_event is not None and cancel_event.is_set():
            terminal_status = SolveStatus.CANCELLED
            break
        elapsed = time.monotonic() - started
        remaining = time_limit_seconds - elapsed
        if remaining <= 0:
            terminal_status = (
                SolveStatus.EXACT_TIMEOUT_WITH_RESULTS
                if unique_solutions
                else SolveStatus.EXACT_NOT_FOUND_TIMEOUT
            )
            break

        solver = cp_model.CpSolver()
        solver.parameters.max_time_in_seconds = remaining
        solver.parameters.num_search_workers = 8
        status = _solve_with_cancellation(solver, model, cancel_event)

        if cancel_event is not None and cancel_event.is_set():
            terminal_status = SolveStatus.CANCELLED
            break
        if status in (cp_model.OPTIMAL, cp_model.FEASIBLE):
            solution = _extract_solution(solver, variables)
            signature = _canonical_signature(solution, problem.targets)
            if signature not in signatures:
                signatures.add(signature)
                unique_solutions.append(solution)
            _exclude_raw_assignment(model, variables, solver)
            continue
        if status == cp_model.INFEASIBLE:
            search_complete = True
            terminal_status = (
                SolveStatus.EXACT_COMPLETE
                if unique_solutions
                else SolveStatus.NO_EXACT_PROVED
            )
            break
        terminal_status = (
            SolveStatus.EXACT_TIMEOUT_WITH_RESULTS
            if unique_solutions
            else SolveStatus.EXACT_NOT_FOUND_TIMEOUT
        )
        break
    else:
        terminal_status = SolveStatus.EXACT_TRUNCATED

    if len(unique_solutions) >= max_solutions and not search_complete:
        terminal_status = SolveStatus.EXACT_TRUNCATED

    return SolveOutcome(
        status=terminal_status,
        exact_solutions=tuple(unique_solutions),
        search_complete=search_complete,
        wall_time_seconds=time.monotonic() - started,
        message=_status_message(terminal_status, len(unique_solutions)),
    )


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

    monitor_thread = threading.Thread(target=monitor, daemon=True)
    monitor_thread.start()
    try:
        return solver.solve(model)
    finally:
        monitor_done.set()
        monitor_thread.join(timeout=0.2)


def _extract_solution(
    solver: cp_model.CpSolver,
    variables: list[list[cp_model.IntVar]],
) -> ExactSolution:
    assignments: list[TargetAssignment] = []
    target_count = len(variables[0]) if variables else 0
    for target_index in range(target_count):
        cell_indices = tuple(
            cell_index
            for cell_index, cell_vars in enumerate(variables)
            if solver.value(cell_vars[target_index])
        )
        assignments.append(TargetAssignment(target_index, cell_indices))
    return ExactSolution(tuple(assignments))


def _exclude_raw_assignment(
    model: cp_model.CpModel,
    variables: list[list[cp_model.IntVar]],
    solver: cp_model.CpSolver,
) -> None:
    matching_literals = []
    for cell_vars in variables:
        for variable in cell_vars:
            matching_literals.append(
                variable if solver.value(variable) else variable.negated()
            )
    model.add_bool_or([literal.negated() for literal in matching_literals])


def _canonical_signature(
    solution: ExactSolution, targets: tuple[int, ...]
) -> tuple:
    by_target_value: dict[int, list[tuple[int, ...]]] = defaultdict(list)
    for assignment in solution.assignments:
        by_target_value[targets[assignment.target_index]].append(
            tuple(sorted(assignment.cell_indices))
        )
    return tuple(
        (target, tuple(sorted(cell_groups)))
        for target, cell_groups in sorted(by_target_value.items())
    )


def _status_message(status: SolveStatus, count: int) -> str:
    messages = {
        SolveStatus.EXACT_COMPLETE: f"已完整搜索并找到 {count} 套精确方案。",
        SolveStatus.EXACT_TRUNCATED: f"已找到 {count} 套精确方案，达到数量上限。",
        SolveStatus.EXACT_TIMEOUT_WITH_RESULTS: (
            f"搜索达到时限，已找到 {count} 套精确方案，可能还有其他方案。"
        ),
        SolveStatus.NO_EXACT_PROVED: "已证明不存在同时满足全部目标的精确方案。",
        SolveStatus.EXACT_NOT_FOUND_TIMEOUT: (
            "搜索达到时限，尚未找到精确方案，不能确认无解。"
        ),
        SolveStatus.CANCELLED: "搜索已取消。",
    }
    return messages[status]

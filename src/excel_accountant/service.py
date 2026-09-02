from __future__ import annotations

import threading
from collections.abc import Callable, Sequence
from dataclasses import dataclass
from pathlib import Path

from .decimal_codec import encode_problem, parse_amount
from .models import (
    OutputArtifact,
    ScaledProblem,
    SolveOutcome,
    SolveStatus,
    TargetAmount,
    WorkbookPreview,
)
from .solver_approximate import solve_approximate
from .solver_exact import solve_exact
from .xlsx_reader import read_workbook_preview
from .xlsx_writer import write_approximate_solution, write_exact_solution

ProgressCallback = Callable[[str, str], None]


class SearchInputError(ValueError):
    """Raised when a complete search request cannot be constructed."""


@dataclass(frozen=True, slots=True)
class SearchRequest:
    source_path: Path
    sheet: str
    range_text: str
    target_values: tuple[str, ...]
    output_directory: Path
    max_exact_solutions: int = 20
    exact_time_limit_seconds: float = 60.0
    max_approximate_solutions: int = 3
    approximate_time_limit_seconds: float = 30.0


@dataclass(frozen=True, slots=True)
class SearchReport:
    request: SearchRequest
    preview: WorkbookPreview
    targets: tuple[TargetAmount, ...]
    scaled_problem: ScaledProblem
    exact_outcome: SolveOutcome
    approximate_outcome: SolveOutcome | None = None
    artifacts: tuple[OutputArtifact, ...] = ()
    messages: tuple[str, ...] = ()

    @property
    def cancelled(self) -> bool:
        return self.exact_outcome.status == SolveStatus.CANCELLED or (
            self.approximate_outcome is not None
            and self.approximate_outcome.status == SolveStatus.CANCELLED
        )


def parse_targets(values: Sequence[str]) -> tuple[TargetAmount, ...]:
    targets: list[TargetAmount] = []
    for value in values:
        raw = value.strip()
        if not raw:
            continue
        targets.append(
            TargetAmount(
                identifier=f"目标{len(targets) + 1}",
                raw_value=raw,
                amount=parse_amount(raw),
            )
        )
    if not targets:
        raise SearchInputError("至少需要输入一个目标金额")
    if len(targets) > 20:
        raise SearchInputError("单次最多支持 20 个目标金额")
    return tuple(targets)


def run_search(
    request: SearchRequest,
    *,
    cancel_event: threading.Event | None = None,
    progress: ProgressCallback | None = None,
) -> SearchReport:
    cancel = cancel_event or threading.Event()
    _notify(progress, "read", "正在读取工作簿并执行安全检查…")
    targets = parse_targets(request.target_values)
    preview = read_workbook_preview(
        request.source_path,
        request.sheet,
        request.range_text,
    )
    if not preview.cells:
        raise SearchInputError("选定范围内没有可用的非零金额")
    problem = encode_problem(
        [cell.amount for cell in preview.cells],
        [target.amount for target in targets],
    )

    _notify(
        progress,
        "exact",
        f"正在 {len(preview.cells)} 条金额中搜索互不重复的精确组合…",
    )
    exact = solve_exact(
        problem,
        max_solutions=request.max_exact_solutions,
        time_limit_seconds=request.exact_time_limit_seconds,
        cancel_event=cancel,
    )
    messages = [exact.message]
    if exact.status == SolveStatus.CANCELLED:
        _notify(progress, "cancelled", exact.message)
        return SearchReport(request, preview, targets, problem, exact, messages=tuple(messages))

    if exact.exact_solutions:
        if not preview.safety.safe_to_write:
            reason_text = "；".join(preview.safety.reasons)
            messages.append(
                f"已找到精确方案，但源文件安全检查未通过，未输出文件：{reason_text}"
            )
            _notify(progress, "unsafe", messages[-1])
        else:
            messages.append(
                f"已找到 {len(exact.exact_solutions)} 套精确方案。"
                "请勾选需要的方案后再点击输出。"
            )
            _notify(progress, "done", messages[-1])
        return SearchReport(
            request=request,
            preview=preview,
            targets=targets,
            scaled_problem=problem,
            exact_outcome=exact,
            messages=tuple(messages),
        )

    _notify(progress, "approximate", "未找到精确方案，正在计算近似备选…")
    approximate = solve_approximate(
        problem,
        max_solutions=request.max_approximate_solutions,
        time_limit_seconds=request.approximate_time_limit_seconds,
        cancel_event=cancel,
    )
    messages.extend(
        (
            "搜索阶段未输出任何文件。",
            approximate.message,
            "近似方案可勾选输出，但实际合计可能不等于目标金额。",
        )
    )
    _notify(progress, "done", " ".join(messages[-2:]))
    return SearchReport(
        request=request,
        preview=preview,
        targets=targets,
        scaled_problem=problem,
        exact_outcome=exact,
        approximate_outcome=approximate,
        messages=tuple(messages),
    )


def _notify(callback: ProgressCallback | None, stage: str, message: str) -> None:
    if callback is not None:
        callback(stage, message)


def export_exact_solutions(
    report: SearchReport,
    solution_indices: Sequence[int],
    output_directory: str | Path,
    *,
    cancel_event: threading.Event | None = None,
    progress: ProgressCallback | None = None,
) -> tuple[OutputArtifact, ...]:
    if not report.preview.safety.safe_to_write:
        reasons = "；".join(report.preview.safety.reasons)
        raise SearchInputError(f"当前工作簿禁止输出：{reasons}")
    indices = tuple(dict.fromkeys(solution_indices))
    if not indices:
        raise SearchInputError("请至少勾选一套精确方案")
    solution_count = len(report.exact_outcome.exact_solutions)
    if any(index < 0 or index >= solution_count for index in indices):
        raise SearchInputError("勾选的方案编号无效，请重新搜索")

    cancel = cancel_event or threading.Event()
    artifacts: list[OutputArtifact] = []
    for position, solution_index in enumerate(indices, start=1):
        if cancel.is_set():
            break
        _notify(
            progress,
            "export",
            f"正在输出并复核 {position}/{len(indices)} 套方案…",
        )
        artifacts.append(
            write_exact_solution(
                report.preview,
                report.targets,
                report.exact_outcome.exact_solutions[solution_index],
                output_directory,
                scheme_number=solution_index + 1,
            )
        )
    _notify(progress, "done", f"已输出 {len(artifacts)} 个并通过复核的 XLSX 文件。")
    return tuple(artifacts)


def export_approximate_solutions(
    report: SearchReport,
    solution_indices: Sequence[int],
    output_directory: str | Path,
    *,
    cancel_event: threading.Event | None = None,
    progress: ProgressCallback | None = None,
) -> tuple[OutputArtifact, ...]:
    if not report.preview.safety.safe_to_write:
        reasons = "；".join(report.preview.safety.reasons)
        raise SearchInputError(f"当前工作簿禁止输出：{reasons}")
    approximate = report.approximate_outcome
    if approximate is None:
        raise SearchInputError("当前搜索结果中没有近似方案")
    indices = tuple(dict.fromkeys(solution_indices))
    if not indices:
        raise SearchInputError("请至少勾选一套近似方案")
    solution_count = len(approximate.approximate_solutions)
    if any(index < 0 or index >= solution_count for index in indices):
        raise SearchInputError("勾选的近似方案编号无效，请重新搜索")

    cancel = cancel_event or threading.Event()
    artifacts: list[OutputArtifact] = []
    for position, solution_index in enumerate(indices, start=1):
        if cancel.is_set():
            break
        _notify(
            progress,
            "export",
            f"正在输出并复核近似方案 {position}/{len(indices)}…",
        )
        artifacts.append(
            write_approximate_solution(
                report.preview,
                report.targets,
                report.scaled_problem,
                approximate.approximate_solutions[solution_index],
                output_directory,
                scheme_number=solution_index + 1,
            )
        )
    _notify(
        progress,
        "done",
        f"已输出 {len(artifacts)} 个近似方案文件，请人工复核差额。",
    )
    return tuple(artifacts)


def export_selected_solutions(
    report: SearchReport,
    solution_indices: Sequence[int],
    output_directory: str | Path,
    *,
    cancel_event: threading.Event | None = None,
    progress: ProgressCallback | None = None,
) -> tuple[OutputArtifact, ...]:
    if report.exact_outcome.exact_solutions:
        return export_exact_solutions(
            report,
            solution_indices,
            output_directory,
            cancel_event=cancel_event,
            progress=progress,
        )
    return export_approximate_solutions(
        report,
        solution_indices,
        output_directory,
        cancel_event=cancel_event,
        progress=progress,
    )

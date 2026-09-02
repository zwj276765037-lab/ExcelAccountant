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
from .xlsx_writer import write_exact_solution

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

    artifacts: list[OutputArtifact] = []
    if exact.exact_solutions:
        if not preview.safety.safe_to_write:
            reason_text = "；".join(preview.safety.reasons)
            messages.append(
                f"已找到精确方案，但源文件安全检查未通过，未输出文件：{reason_text}"
            )
            _notify(progress, "unsafe", messages[-1])
        else:
            _notify(
                progress,
                "write",
                f"正在生成并独立校验 {len(exact.exact_solutions)} 个结果文件…",
            )
            for scheme_number, solution in enumerate(exact.exact_solutions, start=1):
                if cancel.is_set():
                    messages.append("输出过程已取消，已完成的文件保留。")
                    break
                artifacts.append(
                    write_exact_solution(
                        preview,
                        targets,
                        solution,
                        request.output_directory,
                        scheme_number=scheme_number,
                    )
                )
            messages.append(f"已生成 {len(artifacts)} 个通过复核的 XLSX 文件。")
            _notify(progress, "done", messages[-1])
        return SearchReport(
            request=request,
            preview=preview,
            targets=targets,
            scaled_problem=problem,
            exact_outcome=exact,
            artifacts=tuple(artifacts),
            messages=tuple(messages),
        )

    _notify(progress, "approximate", "未找到精确方案，正在计算不写入文件的近似备选…")
    approximate = solve_approximate(
        problem,
        max_solutions=request.max_approximate_solutions,
        time_limit_seconds=request.approximate_time_limit_seconds,
        cancel_event=cancel,
    )
    messages.extend(
        (
            "未输出任何文件。",
            approximate.message,
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

from __future__ import annotations

from dataclasses import dataclass, field
from decimal import Decimal
from enum import StrEnum
from pathlib import Path


class SolveStatus(StrEnum):
    EXACT_COMPLETE = "EXACT_COMPLETE"
    EXACT_TRUNCATED = "EXACT_TRUNCATED"
    EXACT_TIMEOUT_WITH_RESULTS = "EXACT_TIMEOUT_WITH_RESULTS"
    NO_EXACT_PROVED = "NO_EXACT_PROVED"
    EXACT_NOT_FOUND_TIMEOUT = "EXACT_NOT_FOUND_TIMEOUT"
    APPROXIMATE_READY = "APPROXIMATE_READY"
    INVALID_INPUT = "INVALID_INPUT"
    UNSAFE_WORKBOOK = "UNSAFE_WORKBOOK"
    OUTPUT_VERIFICATION_FAILED = "OUTPUT_VERIFICATION_FAILED"
    CANCELLED = "CANCELLED"


@dataclass(frozen=True, slots=True)
class MoneyCell:
    sheet: str
    address: str
    row: int
    column: int
    raw_value: str
    amount: Decimal
    source_type: str = "number"
    hidden: bool = False


@dataclass(frozen=True, slots=True)
class TargetAmount:
    identifier: str
    raw_value: str
    amount: Decimal


@dataclass(frozen=True, slots=True)
class ScaledProblem:
    scale: int
    divisor: int
    amounts: tuple[int, ...]
    targets: tuple[int, ...]

    def restore(self, value: int) -> Decimal:
        return Decimal(value * self.divisor) / Decimal(self.scale)


@dataclass(frozen=True, slots=True)
class TargetAssignment:
    target_index: int
    cell_indices: tuple[int, ...]


@dataclass(frozen=True, slots=True)
class ExactSolution:
    assignments: tuple[TargetAssignment, ...]

    @property
    def used_cell_indices(self) -> tuple[int, ...]:
        return tuple(
            cell_index
            for assignment in self.assignments
            for cell_index in assignment.cell_indices
        )


@dataclass(frozen=True, slots=True)
class ApproximateAssignment:
    target_index: int
    cell_indices: tuple[int, ...]
    actual: int
    difference: int


@dataclass(frozen=True, slots=True)
class ApproximateSolution:
    assignments: tuple[ApproximateAssignment, ...]
    exact_target_count: int
    total_absolute_difference: int
    maximum_absolute_difference: int


@dataclass(frozen=True, slots=True)
class SolveOutcome:
    status: SolveStatus
    exact_solutions: tuple[ExactSolution, ...] = ()
    approximate_solutions: tuple[ApproximateSolution, ...] = ()
    search_complete: bool = False
    wall_time_seconds: float = 0.0
    message: str = ""


@dataclass(frozen=True, slots=True)
class SkippedCell:
    sheet: str
    address: str
    raw_value: str
    reason: str


@dataclass(frozen=True, slots=True)
class WorkbookSafetyReport:
    safe_to_write: bool
    reasons: tuple[str, ...] = ()


@dataclass(frozen=True, slots=True)
class WorkbookPreview:
    path: Path
    sheet: str
    range_text: str
    cells: tuple[MoneyCell, ...]
    skipped: tuple[SkippedCell, ...]
    formula_count: int = 0
    zero_count: int = 0
    hidden_count: int = 0
    safety: WorkbookSafetyReport = field(
        default_factory=lambda: WorkbookSafetyReport(True)
    )


@dataclass(frozen=True, slots=True)
class OutputArtifact:
    path: Path
    result_sheet: str
    scheme_number: int

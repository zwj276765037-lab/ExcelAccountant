from __future__ import annotations

import threading
from pathlib import Path

import pytest
from openpyxl import Workbook

from excel_accountant.models import SolveStatus
from excel_accountant.service import (
    SearchInputError,
    SearchRequest,
    export_exact_solutions,
    parse_targets,
    run_search,
)


def _workbook(path: Path, values: tuple[object, ...]) -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "流水"
    sheet["E1"] = "金额"
    for row, value in enumerate(values, start=2):
        sheet.cell(row, 5, value)
    workbook.save(path)
    workbook.close()


def _request(source: Path, output: Path, *targets: str) -> SearchRequest:
    return SearchRequest(
        source_path=source,
        sheet="流水",
        range_text="E",
        target_values=targets,
        output_directory=output,
        max_exact_solutions=3,
        exact_time_limit_seconds=5,
        max_approximate_solutions=2,
        approximate_time_limit_seconds=5,
    )


def test_parse_targets_ignores_blank_rows_and_preserves_decimal() -> None:
    targets = parse_targets((" 10.00 ", "", "-2.345"))
    assert [item.identifier for item in targets] == ["目标1", "目标2"]
    assert [str(item.amount) for item in targets] == ["10.00", "-2.345"]


def test_parse_targets_rejects_empty_input() -> None:
    with pytest.raises(SearchInputError, match="至少"):
        parse_targets(("", "  "))


def test_exact_search_waits_for_selection_before_export(tmp_path: Path) -> None:
    source = tmp_path / "source.xlsx"
    _workbook(source, ("1.10", "2.20", "3.30", "4.40"))
    progress: list[str] = []
    report = run_search(
        _request(source, tmp_path / "output", "3.30", "7.70"),
        progress=lambda stage, _message: progress.append(stage),
    )

    assert report.exact_outcome.exact_solutions
    assert report.approximate_outcome is None
    assert not report.artifacts
    assert not (tmp_path / "output").exists()
    assert "exact" in progress and progress[-1] == "done"

    artifacts = export_exact_solutions(
        report,
        (1,),
        tmp_path / "selected-output",
    )
    assert len(artifacts) == 1
    assert artifacts[0].scheme_number == 2
    assert artifacts[0].path.exists()


def test_export_rejects_empty_selection(tmp_path: Path) -> None:
    source = tmp_path / "source.xlsx"
    _workbook(source, ("1.00", "2.00"))
    report = run_search(_request(source, tmp_path / "output", "3.00"))

    with pytest.raises(SearchInputError, match="勾选"):
        export_exact_solutions(report, (), tmp_path / "selected-output")


def test_no_exact_solution_returns_approximate_without_files(tmp_path: Path) -> None:
    source = tmp_path / "source.xlsx"
    _workbook(source, ("1.00", "2.00"))
    output = tmp_path / "output"
    report = run_search(_request(source, output, "10.00"))

    assert report.exact_outcome.status == SolveStatus.NO_EXACT_PROVED
    assert report.approximate_outcome is not None
    assert report.approximate_outcome.approximate_solutions
    assert not report.artifacts
    assert not output.exists()


def test_cancelled_request_does_not_write_files(tmp_path: Path) -> None:
    source = tmp_path / "source.xlsx"
    _workbook(source, ("1.00", "2.00"))
    cancel = threading.Event()
    cancel.set()
    report = run_search(
        _request(source, tmp_path / "output", "3.00"),
        cancel_event=cancel,
    )

    assert report.cancelled
    assert not report.artifacts

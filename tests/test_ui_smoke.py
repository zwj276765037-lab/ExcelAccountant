from __future__ import annotations

import os
from pathlib import Path

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from openpyxl import Workbook
from PySide6.QtCore import Qt
from PySide6.QtWidgets import QApplication

from excel_accountant.main_window import MainWindow
from excel_accountant.service import SearchRequest, run_search


def test_main_window_can_be_created_offscreen() -> None:
    application = QApplication.instance() or QApplication([])
    window = MainWindow()
    try:
        assert "ExcelAccountant" in window.windowTitle()
        assert window.range_edit.text() == "E"
        assert window.solution_count.value() == 20
        assert window.search_button.isEnabled()
        assert not window.export_button.isEnabled()
    finally:
        window.close()
        application.processEvents()


def test_exact_schemes_render_with_one_checkbox_per_scheme(tmp_path: Path) -> None:
    source = tmp_path / "source.xlsx"
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "流水"
    worksheet["E1"] = "金额"
    for row, value in enumerate(("1.10", "2.20", "3.30", "4.40"), start=2):
        worksheet.cell(row, 5, value)
    workbook.save(source)
    workbook.close()
    report = run_search(
        SearchRequest(
            source,
            "流水",
            "E",
            ("3.30", "7.70"),
            tmp_path / "output",
            max_exact_solutions=2,
            exact_time_limit_seconds=5,
        )
    )

    application = QApplication.instance() or QApplication([])
    window = MainWindow()
    try:
        window._render_report(report)
        window._set_running(False)
        checkbox_rows = [
            row
            for row in range(window.result_table.rowCount())
            if window.result_table.item(row, 0).flags()
            & Qt.ItemFlag.ItemIsUserCheckable
        ]
        assert len(checkbox_rows) == len(report.exact_outcome.exact_solutions)
        assert all(
            window.result_table.item(row, 0).checkState()
            == Qt.CheckState.Unchecked
            for row in checkbox_rows
        )
        assert window.export_button.isEnabled()
        assert not (tmp_path / "output").exists()
    finally:
        window.close()
        application.processEvents()

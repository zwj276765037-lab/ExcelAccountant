from __future__ import annotations

import os

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtWidgets import QApplication

from excel_accountant.main_window import MainWindow


def test_main_window_can_be_created_offscreen() -> None:
    application = QApplication.instance() or QApplication([])
    window = MainWindow()
    try:
        assert "ExcelAccountant" in window.windowTitle()
        assert window.range_edit.text() == "E"
        assert window.solution_count.value() == 20
        assert window.search_button.isEnabled()
    finally:
        window.close()
        application.processEvents()

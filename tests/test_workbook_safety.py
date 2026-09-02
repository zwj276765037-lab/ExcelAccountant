from __future__ import annotations

from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

from openpyxl import Workbook

from excel_accountant.workbook_safety import inspect_workbook_safety


def _simple_workbook(path: Path, protected: bool = False) -> None:
    workbook = Workbook()
    workbook.active["A1"] = 1
    workbook.active.protection.sheet = protected
    workbook.save(path)


def _append_part(path: Path, name: str) -> None:
    with ZipFile(path, "a", ZIP_DEFLATED) as archive:
        archive.writestr(name, b"<xml/>")


def test_plain_xlsx_is_safe(tmp_path: Path) -> None:
    path = tmp_path / "plain.xlsx"
    _simple_workbook(path)
    report = inspect_workbook_safety(path)
    assert report.safe_to_write is True
    assert not report.reasons


def test_drawings_are_rejected_for_output(tmp_path: Path) -> None:
    path = tmp_path / "drawing.xlsx"
    _simple_workbook(path)
    _append_part(path, "xl/drawings/drawing1.xml")
    report = inspect_workbook_safety(path)
    assert report.safe_to_write is False
    assert "包含绘图、图片或形状" in report.reasons


def test_protected_sheet_is_rejected_for_output(tmp_path: Path) -> None:
    path = tmp_path / "protected.xlsx"
    _simple_workbook(path, protected=True)
    report = inspect_workbook_safety(path)
    assert report.safe_to_write is False
    assert report.reasons == ("工作表受到保护：Sheet",)


def test_missing_workbook_security_record_is_safe(tmp_path: Path) -> None:
    path = tmp_path / "plain.xlsx"
    _simple_workbook(path)
    report = inspect_workbook_safety(path)
    assert report.safe_to_write is True

from __future__ import annotations

from datetime import date
from decimal import Decimal
from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import ZIP_DEFLATED, ZipFile

import pytest
from openpyxl import Workbook

from excel_accountant.xlsx_reader import (
    WorkbookReadError,
    list_sheet_names,
    parse_column_range,
    read_workbook_preview,
)

MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"


def _make_workbook(path: Path) -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "流水"
    sheet["E1"] = "交易金额"
    sheet["E2"] = 1.25
    sheet["E3"] = "2.5000"
    sheet["E4"] = 0
    sheet["E5"] = date(2026, 9, 2)
    sheet["E6"] = "=1+2"
    sheet["E7"] = -3.75
    sheet.row_dimensions[7].hidden = True
    workbook.create_sheet("其他")
    workbook.save(path)


def _patch_numeric_token(path: Path, address: str, value: str) -> None:
    with ZipFile(path, "r") as archive:
        entries = {name: archive.read(name) for name in archive.namelist()}
    sheet_name = "xl/worksheets/sheet1.xml"
    root = ET.fromstring(entries[sheet_name])
    for cell in root.iter(f"{{{MAIN_NS}}}c"):
        if cell.attrib.get("r") == address:
            value_node = cell.find(f"{{{MAIN_NS}}}v")
            assert value_node is not None
            value_node.text = value
            break
    entries[sheet_name] = ET.tostring(root, encoding="utf-8", xml_declaration=True)
    with ZipFile(path, "w", ZIP_DEFLATED) as archive:
        for name, content in entries.items():
            archive.writestr(name, content)


def test_parse_column_range_supports_three_input_forms() -> None:
    assert parse_column_range("E").column == 5
    assert parse_column_range("5").column == 5
    parsed = parse_column_range("E2:E500")
    assert (parsed.column, parsed.start_row, parsed.end_row) == (5, 2, 500)
    assert parse_column_range("E:E").end_row is None


def test_parse_column_range_rejects_multiple_columns() -> None:
    with pytest.raises(WorkbookReadError):
        parse_column_range("E2:F10")


def test_list_sheets_and_read_exact_preview(tmp_path: Path) -> None:
    path = tmp_path / "input.xlsx"
    _make_workbook(path)
    _patch_numeric_token(path, "E2", "1.250000000000001")

    assert list_sheet_names(path) == ("流水", "其他")
    preview = read_workbook_preview(path, "流水", "E")

    assert preview.range_text == "E1:E7"
    assert [cell.address for cell in preview.cells] == ["E2", "E3", "E7"]
    assert preview.cells[0].amount == Decimal("1.250000000000001")
    assert preview.cells[0].raw_value == "1.250000000000001"
    assert preview.cells[1].amount == Decimal("2.5000")
    assert preview.cells[1].source_type == "text"
    assert preview.cells[2].amount == Decimal("-3.75")
    assert preview.cells[2].hidden is True
    assert preview.formula_count == 1
    assert preview.zero_count == 1
    assert preview.hidden_count == 1
    assert {item.reason for item in preview.skipped} == {
        "表头",
        "零值",
        "日期单元格",
        "公式单元格",
    }
    assert preview.safety.safe_to_write is True


def test_explicit_range_excludes_header_and_tail(tmp_path: Path) -> None:
    path = tmp_path / "input.xlsx"
    _make_workbook(path)
    preview = read_workbook_preview(path, "流水", "E2:E3")
    assert preview.range_text == "E2:E3"
    assert [cell.address for cell in preview.cells] == ["E2", "E3"]
    assert not preview.skipped


def test_missing_sheet_is_reported(tmp_path: Path) -> None:
    path = tmp_path / "input.xlsx"
    _make_workbook(path)
    with pytest.raises(WorkbookReadError, match="工作表不存在"):
        read_workbook_preview(path, "不存在", "E")

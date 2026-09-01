from pathlib import Path

import pytest
from openpyxl import Workbook

from app.excel.excel_reader import ExcelReader, ExcelReaderError


def _make_workbook(path: Path) -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Data"
    sheet["B5"] = "first"
    sheet["B6"] = 42
    sheet["B7"] = ""
    sheet["B8"] = "must not be reached"
    workbook.create_sheet("Other")
    workbook.save(path)


def test_reads_down_one_column_and_stops_at_first_empty_cell(tmp_path: Path) -> None:
    workbook_path = tmp_path / "relay.xlsx"
    _make_workbook(workbook_path)
    reader = ExcelReader()

    reader.load(workbook_path)
    assert reader.sheet_names == ["Data", "Other"]
    reader.select_sheet("Data")
    reader.set_start_cell("b5")

    assert reader.current_cell_address == "B5"
    assert reader.current_text == "first"
    assert reader.estimate_total_items() == 2
    assert reader.advance() == "B6"
    assert reader.current_text == "42"
    assert reader.advance() == "B7"
    assert reader.current_text is None


@pytest.mark.parametrize("address", ["", "B0", "A1:B2", "XFE1", "A1048577"])
def test_rejects_invalid_or_out_of_range_cell_addresses(
    tmp_path: Path, address: str
) -> None:
    workbook_path = tmp_path / "relay.xlsx"
    _make_workbook(workbook_path)
    reader = ExcelReader()
    reader.load(workbook_path)
    reader.select_sheet("Data")

    with pytest.raises(ExcelReaderError):
        reader.set_start_cell(address)


def test_rejects_non_xlsx_file(tmp_path: Path) -> None:
    path = tmp_path / "relay.xls"
    path.write_bytes(b"not a workbook")
    with pytest.raises(ExcelReaderError, match="xlsx"):
        ExcelReader().load(path)

"""Single-column, row-by-row Excel reader for CellRelay."""

from __future__ import annotations

import re
from pathlib import Path
from typing import Any

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.workbook.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet

_CELL_RE = re.compile(r"^([A-Za-z]{1,3})([1-9][0-9]*)$")
_MAX_EXCEL_COLUMN = 16_384
_MAX_EXCEL_ROW = 1_048_576


class ExcelReaderError(RuntimeError):
    """Raised for invalid workbook operations or coordinates."""


class ExcelReader:
    """Read values down one column without automating the Excel application."""

    def __init__(self) -> None:
        self._workbook: Workbook | None = None
        self._worksheet: Worksheet | None = None
        self._path: Path | None = None
        self._column = 1
        self._row = 1
        self._start_row = 1
        self._total_items_cache: int | None = None

    @property
    def path(self) -> Path | None:
        return self._path

    @property
    def selected_sheet(self) -> str | None:
        return self._worksheet.title if self._worksheet is not None else None

    @property
    def sheet_names(self) -> list[str]:
        self._require_workbook()
        assert self._workbook is not None
        return list(self._workbook.sheetnames)

    @property
    def current_cell_address(self) -> str:
        return f"{get_column_letter(self._column)}{self._row}"

    @property
    def current_value(self) -> Any:
        worksheet = self._require_worksheet()
        return worksheet.cell(row=self._row, column=self._column).value

    @property
    def current_text(self) -> str | None:
        """Return the current value as text, or None for an empty cell."""
        value = self.current_value
        if value is None or value == "":
            return None
        return str(value)

    def load(self, file_path: str | Path) -> None:
        """Open an xlsx workbook using cached formula results when available."""
        path = Path(file_path).expanduser().resolve()
        if path.suffix.lower() != ".xlsx":
            raise ExcelReaderError(".xlsx 파일만 지원합니다.")
        if not path.is_file():
            raise ExcelReaderError(f"Excel 파일을 찾을 수 없습니다: {path}")
        self.close()
        try:
            self._workbook = load_workbook(
                filename=path,
                read_only=False,
                data_only=True,
            )
        except Exception as exc:
            raise ExcelReaderError(f"Excel 파일을 열지 못했습니다: {exc}") from exc
        self._path = path
        self._worksheet = None
        self._total_items_cache = None

    def select_sheet(self, sheet_name: str) -> None:
        workbook = self._require_workbook()
        if sheet_name not in workbook.sheetnames:
            raise ExcelReaderError(f"Sheet를 찾을 수 없습니다: {sheet_name}")
        self._worksheet = workbook[sheet_name]
        self._total_items_cache = None

    def set_start_cell(self, address: str) -> None:
        """Set the initial coordinate while preserving its column thereafter."""
        self._require_worksheet()
        normalized = address.strip().upper()
        match = _CELL_RE.fullmatch(normalized)
        if match is None:
            raise ExcelReaderError(f"올바른 셀 주소가 아닙니다: {address}")
        column_letters, row_text = match.groups()
        column = 0
        for character in column_letters:
            column = column * 26 + (ord(character) - ord("A") + 1)
        row = int(row_text)
        if column > _MAX_EXCEL_COLUMN or row > _MAX_EXCEL_ROW:
            raise ExcelReaderError(f"Excel 범위를 벗어난 셀입니다: {address}")
        self._column = column
        self._row = row
        self._start_row = row
        self._total_items_cache = None

    def advance(self) -> str:
        """Move exactly one row only after the controller records completion."""
        if self._row >= _MAX_EXCEL_ROW:
            raise ExcelReaderError("Excel의 마지막 행에 도달했습니다.")
        self._row += 1
        return self.current_cell_address

    def estimate_total_items(self) -> int:
        """Count contiguous non-empty cells from the configured start cell."""
        if self._total_items_cache is not None:
            return self._total_items_cache
        worksheet = self._require_worksheet()
        count = 0
        for row in range(self._start_row, min(worksheet.max_row, _MAX_EXCEL_ROW) + 1):
            value = worksheet.cell(row=row, column=self._column).value
            if value is None or value == "":
                break
            count += 1
        self._total_items_cache = count
        return count

    def close(self) -> None:
        if self._workbook is not None:
            self._workbook.close()
        self._workbook = None
        self._worksheet = None
        self._path = None
        self._total_items_cache = None

    def _require_workbook(self) -> Workbook:
        if self._workbook is None:
            raise ExcelReaderError("먼저 Excel 파일을 선택하세요.")
        return self._workbook

    def _require_worksheet(self) -> Worksheet:
        if self._worksheet is None:
            raise ExcelReaderError("먼저 Sheet를 선택하세요.")
        return self._worksheet

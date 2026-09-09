"""A disposable native-Excel check for deployed executables."""

from __future__ import annotations

import logging
from pathlib import Path
from tempfile import TemporaryDirectory

from app.excel.com_excel import excel_application
from app.excel.excel_reader import ExcelReader
from app.excel.result_writer import HEADER, ExcelResultWriter, signature


def run_excel_smoke_test() -> int:
    logger = logging.getLogger(__name__)
    reader = ExcelReader(backend="com")
    try:
        with TemporaryDirectory(prefix="CellRelay-excel-smoke-") as directory:
            legacy = Path(directory) / "users.xls"
            with excel_application() as application:
                try:
                    book = application.Workbooks.Add(-4167)
                    try:
                        sheet = book.Worksheets.Item(1)
                        try:
                            sheet.Name = "Data"
                            sheet.Range("B1:B2").Value2 = (("first",), ("second",))
                        finally:
                            sheet = None
                        book.SaveAs(str(legacy), FileFormat=56)
                    finally:
                        book.Close(SaveChanges=False)
                        book = None
                finally:
                    application = None
            # A real binary workbook exercises the non-OOXML path without
            # requiring a user's protected file or changing its extension.
            path = legacy
            original = path.read_bytes()
            reader.load(path)
            reader.select_sheet("Data")
            reader.set_start_cell("B1")
            if reader.current_text != "first" or reader.estimate_total_items() != 2:
                raise RuntimeError("Excel COM input verification failed")
            _, backup = ExcelResultWriter().write_error(
                path, "Data", "B2", "SMOKE_TEST", signature(path), backend="com"
            )
            if Path(backup).read_bytes() != original:
                raise RuntimeError("Excel COM backup verification failed")
            reader.advance()
            if reader.current_text != "second":
                raise RuntimeError("Excel COM input snapshot changed after saving")
            reader.load(path)
            reader.select_sheet("Data")
            reader.set_start_cell("C1")
            if reader.current_text != HEADER:
                raise RuntimeError("Excel COM error column was not saved")
            reader.advance()
            if "SMOKE_TEST" not in (reader.current_text or ""):
                raise RuntimeError("Excel COM error entry was not saved")
        logger.info("Packaged Excel COM smoke test passed")
        return 0
    except Exception:
        logger.exception("Packaged Excel COM smoke test failed")
        return 1
    finally:
        reader.close()

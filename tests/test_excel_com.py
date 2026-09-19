"""COM routing/failure tests and real Excel regressions when Excel is installed."""

import gc
import traceback
import weakref
from io import BytesIO
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import Mock, PropertyMock
from zipfile import ZipFile, is_zipfile

import pytest
from openpyxl import Workbook
from PySide6.QtCore import Qt, QThread

from app.excel import com_excel
from app.excel.excel_reader import ExcelReader, ExcelReaderError
from app.excel.result_writer import HEADER, ExcelResultWriter, ExcelWorker, signature


def test_embedded_theme_zip_is_not_mistaken_for_an_ooxml_workbook(
    tmp_path, monkeypatch
):
    theme = BytesIO()
    with ZipFile(theme, "w") as archive:
        archive.writestr("theme/theme/theme1.xml", "<theme />")
    path = tmp_path / "legacy.xlsx"
    original = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1" + theme.getvalue()
    path.write_bytes(original)
    assert is_zipfile(path)  # Python recognizes the embedded ZIP at the end.
    assert not com_excel.is_ooxml_workbook(path)
    save = Mock(side_effect=RuntimeError("COM selected"))
    monkeypatch.setattr(com_excel, "save_error_copy", save)
    with pytest.raises(RuntimeError, match="COM selected"):
        ExcelResultWriter().write_error(path, "Data", "B1", "error", signature(path))
    save.assert_called_once()
    assert path.read_bytes() == original


def test_non_zip_reader_and_writer_never_use_zip_and_keep_snapshot(
    tmp_path, monkeypatch
):
    path = tmp_path / "protected.xlsx"
    original = b"opaque file readable only by the mocked desktop Excel"
    path.write_bytes(original)
    monkeypatch.setattr(com_excel, "sheet_names", lambda path: ["Data", "Other"])
    read = Mock(return_value=("first", 42, 0, False))
    monkeypatch.setattr(com_excel, "read_column", read)

    def reject_zip(*args, **kwargs):
        pytest.fail("The COM path must not open the file as a ZIP")

    monkeypatch.setattr("app.excel.excel_reader.load_workbook", reject_zip)
    monkeypatch.setattr("app.excel.result_writer.ZipFile", reject_zip)

    def save_copy(source, destination, sheet, row, header, entry):
        assert source == path and sheet == "Data" and row == 5
        assert header == HEADER and "NOT_FOUND" in entry
        destination.write_bytes(b"Excel-generated opaque result")

    monkeypatch.setattr(com_excel, "save_error_copy", save_copy)
    reader = ExcelReader(backend="com")
    reader.load(path)
    assert reader.backend == "com" and reader.sheet_names == ["Data", "Other"]
    reader.select_sheet("Data")
    reader.set_start_cell("b5")
    assert reader.estimate_total_items() == 4 and reader.current_text == "first"
    sig, backup = ExcelResultWriter().write_error(
        path, "Data", "B5", "NOT_FOUND", signature(path), backend=reader.backend
    )
    assert sig == signature(path)
    assert Path(backup).read_bytes() == original
    assert reader.advance() == "B6" and reader.current_text == "42"
    reader.advance()
    assert reader.current_text == "0"
    reader.advance()
    assert reader.current_text == "False"
    reader.advance()
    assert reader.current_text is None
    read.assert_called_once_with(path, "Data", 2, 5)
    reader.close()
    assert (
        reader.path is None and reader.selected_sheet is None and reader.backend is None
    )


def test_auto_fallback_only_when_com_is_unavailable(tmp_path, monkeypatch):
    monkeypatch.setattr("app.excel.excel_reader.sys.platform", "win32")
    path = tmp_path / "users.xlsx"
    book = Workbook()
    book.active["A1"] = "first"
    book.save(path)
    book.close()
    monkeypatch.setattr(
        com_excel,
        "sheet_names",
        Mock(side_effect=com_excel.ExcelComUnavailable("no Excel")),
    )
    reader = ExcelReader()
    reader.load(path)
    assert reader.backend == "ooxml"
    reader.close()
    path.write_bytes(b"not an OOXML ZIP")
    with pytest.raises(ExcelReaderError, match="COM.*no Excel"):
        reader.load(path)
    assert reader.path is None and reader.backend is None
    monkeypatch.setattr(
        com_excel, "sheet_names", Mock(side_effect=RuntimeError("access denied"))
    )
    with pytest.raises(ExcelReaderError, match="access denied"):
        reader.load(path)


def test_com_read_refuses_changed_source(tmp_path, monkeypatch):
    path = tmp_path / "users.xls"
    path.write_bytes(b"original")
    monkeypatch.setattr(com_excel, "sheet_names", lambda path: ["Data"])
    read = Mock(return_value=("first",))
    monkeypatch.setattr(com_excel, "read_column", read)
    reader = ExcelReader(backend="com")
    reader.load(path)
    reader.select_sheet("Data")
    path.write_bytes(b"changed workbook")
    with pytest.raises(ExcelReaderError, match="외부에서 변경"):
        reader.set_start_cell("B5")
    read.assert_not_called()


@pytest.mark.parametrize("failure", ["save", "external_change", "replace"])
def test_com_save_failures_preserve_source(tmp_path, monkeypatch, failure):
    path = tmp_path / "users.xlsb"
    original = b"original binary workbook"
    path.write_bytes(original)

    def save_copy(source, destination, *args):
        destination.write_bytes(b"partial or verified Excel copy")
        if failure == "save":
            raise RuntimeError("Excel save/verification failed")
        if failure == "external_change":
            source.write_bytes(b"external edit")

    monkeypatch.setattr(com_excel, "save_error_copy", save_copy)
    if failure == "replace":
        monkeypatch.setattr(
            "app.excel.result_writer.os.replace",
            Mock(side_effect=PermissionError("locked")),
        )
    with pytest.raises((RuntimeError, PermissionError)):
        ExcelResultWriter().write_error(
            path, "Data", "B5", "NOT_FOUND", signature(path), backend="com"
        )
    expected = b"external edit" if failure == "external_change" else original
    assert path.read_bytes() == expected
    assert not list(tmp_path.glob(".cellrelay-*"))
    backups = list(tmp_path.glob("*.cellrelay-backup-*.xlsb"))
    if failure == "replace":
        assert len(backups) == 1 and backups[0].read_bytes() == original
    else:
        assert not backups


@pytest.mark.parametrize("failure", ["startup", "open", "read"])
def test_com_failure_always_releases_its_own_apartment(tmp_path, monkeypatch, failure):
    pythoncom = Mock(COINIT_APARTMENTTHREADED=2)
    application = Mock()
    client = Mock()
    client.DispatchEx.return_value = application
    book = application.Workbooks.Open.return_value
    worksheet = book.Worksheets.Item.return_value
    worksheet.Columns.Count = 16384
    worksheet.Rows.Count = 1048576
    worksheet.UsedRange.Row = 1
    worksheet.UsedRange.Rows.Count = 5
    if failure == "startup":
        client.DispatchEx.side_effect = RuntimeError("startup failed")
    elif failure == "open":
        application.Workbooks.Open.side_effect = RuntimeError("open failed")
    else:
        worksheet.Range.side_effect = RuntimeError("read failed")
    monkeypatch.setattr(com_excel, "_load_runtime", lambda: (pythoncom, client))
    with pytest.raises(RuntimeError):
        com_excel.read_column(tmp_path / "users.xlsx", "Data", 2, 1)
    pythoncom.CoInitializeEx.assert_called_once_with(2)
    pythoncom.CoUninitialize.assert_called_once_with()
    client.DispatchEx.assert_called_once_with("Excel.Application")
    if failure != "startup":
        application.Quit.assert_called_once_with()
        assert application.Visible is False and application.EnableEvents is False
        assert application.AutomationSecurity == 3
        options = application.Workbooks.Open.call_args.kwargs
        assert options["ReadOnly"] is True and options["UpdateLinks"] == 0
        assert options["Password"] == "" and options["Notify"] is False
    if failure == "read":
        book.Close.assert_called_once_with(SaveChanges=False)


class RejectedComCall(Exception):
    def __init__(self, hresult):
        super().__init__("Excel rejected the call")
        self.hresult = hresult


@pytest.mark.parametrize("hresult", [-2147418111, -2147417846, 0x80010001, 0x8001010A])
def test_com_retries_explicit_rejection_and_logs_stage(monkeypatch, caplog, hresult):
    sleep = Mock()
    monkeypatch.setattr(com_excel.time, "sleep", sleep)
    call = Mock(side_effect=[RejectedComCall(hresult), "done"])
    with caplog.at_level("INFO", logger=com_excel.__name__):
        assert com_excel._com_call("Workbook.SaveCopyAs", call) == "done"
    assert call.call_count == 2
    sleep.assert_called_once_with(0.1)
    assert "started operation=Workbook.SaveCopyAs" in caplog.text
    assert f"hresult=0x{hresult & 0xFFFFFFFF:08X}" in caplog.text
    assert "completed operation=Workbook.SaveCopyAs attempts=2 elapsed=" in caplog.text


@pytest.mark.parametrize(
    "error",
    [
        RuntimeError("unknown save outcome"),
        RejectedComCall(0x80020009),
        RejectedComCall(0x80010108),
    ],
)
def test_com_does_not_retry_ambiguous_failures(monkeypatch, error):
    sleep = Mock()
    monkeypatch.setattr(com_excel.time, "sleep", sleep)
    call = Mock(side_effect=error)
    with pytest.raises(type(error)) as caught:
        com_excel._com_call("Workbook.SaveCopyAs", call)
    assert caught.value is error
    call.assert_called_once_with()
    sleep.assert_not_called()


def test_com_rejection_retries_are_bounded_and_preserve_error(monkeypatch):
    sleep = Mock()
    monkeypatch.setattr(com_excel.time, "sleep", sleep)
    monkeypatch.setattr(com_excel.time, "monotonic", lambda: 0.0)
    error = RejectedComCall(0x8001010A)
    call = Mock(side_effect=error)
    with pytest.raises(RejectedComCall) as caught:
        com_excel._com_call("Workbook.Close", call)
    assert caught.value is error
    assert call.call_count == com_excel._COM_RETRY_MAX_ATTEMPTS
    assert sum(args.args[0] for args in sleep.call_args_list) == pytest.approx(2.5)


def test_com_rejection_retry_stops_at_elapsed_deadline(monkeypatch):
    sleep = Mock()
    monkeypatch.setattr(com_excel.time, "sleep", sleep)
    monkeypatch.setattr(com_excel.time, "monotonic", Mock(side_effect=[0, 2.9, 3, 3]))
    error = RejectedComCall(0x80010001)
    call = Mock(side_effect=error)
    with pytest.raises(RejectedComCall) as caught:
        com_excel._com_call("Workbook.Close", call)
    assert caught.value is error
    call.assert_called_once_with()
    assert sleep.call_args.args[0] == pytest.approx(0.1)


@pytest.fixture
def mocked_com_runtime(monkeypatch):
    pythoncom = Mock(COINIT_APARTMENTTHREADED=2)
    application = Mock()
    client = Mock()
    client.DispatchEx.return_value = application
    monkeypatch.setattr(com_excel, "_load_runtime", lambda: (pythoncom, client))
    monkeypatch.setattr(com_excel.time, "sleep", Mock())
    return pythoncom, client, application


@pytest.mark.parametrize("failed_property", ["NumberFormat", "Value2"])
def test_com_failure_releases_callback_proxy_and_preserves_exception(
    monkeypatch, failed_property
):
    error = RejectedComCall(0x80010001)
    monkeypatch.setattr(com_excel, "_COM_RETRY_MAX_ATTEMPTS", 1)

    class FakeCell:
        def __setattr__(self, name, value):
            if name == failed_property:
                raise error
            object.__setattr__(self, name, value)

    references = []

    def get_range(address):
        if address == "A1:A1":
            return SimpleNamespace(Value2="input")
        cell = FakeCell()
        references.append(weakref.ref(cell))
        return cell

    worksheet = SimpleNamespace(
        UsedRange=SimpleNamespace(Column=1, Columns=SimpleNamespace(Count=1)),
        Columns=SimpleNamespace(Count=10),
        Rows=SimpleNamespace(Count=10),
        Range=get_range,
    )
    with pytest.raises(RejectedComCall) as caught:
        com_excel._write_error_cells(worksheet, 1, "header", "entry")
    gc.collect()

    # The exception is deliberately retained, as it is during workbook/Excel
    # cleanup. Neither callback defaults nor traceback locals may keep the cell.
    assert len(references) == 1 and references[0]() is None
    assert caught.value is error and error.hresult == 0x80010001
    frames = list(traceback.walk_tb(error.__traceback__))
    names = [frame.f_code.co_name for frame, _ in frames]
    assert "_write_error_cells" in names and "_com_call" in names
    assert "__setattr__" in names
    call_frame = next(
        frame for frame, _ in frames if frame.f_code.co_name == "_com_call"
    )
    assert call_frame.f_locals["operation"] == f"Range.{failed_property}.set"
    assert call_frame.f_locals["callback"] is None


def test_com_retries_rejected_application_property_set(mocked_com_runtime, monkeypatch):
    pythoncom, client, application = mocked_com_runtime
    visible = PropertyMock(side_effect=[RejectedComCall(0x80010001), None])
    monkeypatch.setattr(type(application), "Visible", visible, raising=False)
    with com_excel.excel_application() as opened:
        assert opened is application
    assert visible.call_count == 2
    application.Quit.assert_called_once_with()
    pythoncom.CoUninitialize.assert_called_once_with()
    client.DispatchEx.assert_called_once_with("Excel.Application")


def test_com_retries_rejected_open_without_restarting_excel(
    mocked_com_runtime, tmp_path
):
    pythoncom, client, application = mocked_com_runtime
    book = Mock()
    application.Workbooks.Open.side_effect = [RejectedComCall(0x8001010A), book]
    with (
        com_excel.excel_application() as opened,
        com_excel.open_workbook(opened, tmp_path / "users.xls", read_only=True),
    ):
        pass
    assert application.Workbooks.Open.call_count == 2
    client.DispatchEx.assert_called_once_with("Excel.Application")
    book.Close.assert_called_once_with(SaveChanges=False)
    pythoncom.CoUninitialize.assert_called_once_with()


def test_com_retries_only_rejected_save_call(mocked_com_runtime, tmp_path, monkeypatch):
    _, client, application = mocked_com_runtime
    book = application.Workbooks.Open.return_value
    book.SaveCopyAs.side_effect = [RejectedComCall(0x8001010A), None]
    book.Worksheets.Item.return_value.Range.return_value.Value2 = "entry"
    edit = Mock(return_value={"B5": "entry"})
    monkeypatch.setattr(com_excel, "_write_error_cells", edit)
    com_excel.save_error_copy(
        tmp_path / "users.xls", tmp_path / "copy.xls", "Data", 5, "header", "entry"
    )
    assert book.SaveCopyAs.call_count == 2
    edit.assert_called_once()
    assert application.Workbooks.Open.call_count == 2  # Source and verification copy.
    client.DispatchEx.assert_called_once_with("Excel.Application")


def test_com_cleanup_retries_preserve_original_failure(
    mocked_com_runtime, tmp_path, caplog
):
    pythoncom, _, application = mocked_com_runtime
    book = application.Workbooks.Open.return_value
    book.Close.side_effect = RejectedComCall(0x8001010A)
    application.Quit.side_effect = RejectedComCall(0x80010001)
    error = RuntimeError("original write failure")
    with (
        pytest.raises(RuntimeError) as caught,
        com_excel.excel_application() as opened,
        com_excel.open_workbook(opened, tmp_path / "users.xls", read_only=True),
    ):
        raise error
    assert caught.value is error
    assert book.Close.call_count == com_excel._COM_RETRY_MAX_ATTEMPTS
    assert application.Quit.call_count == com_excel._COM_RETRY_MAX_ATTEMPTS
    assert "Could not close the CellRelay workbook" in caplog.text
    assert "Could not quit the CellRelay Excel instance" in caplog.text
    pythoncom.CoUninitialize.assert_called_once_with()


@pytest.fixture(scope="module")
def native_workbooks(tmp_path_factory):
    """Create only disposable fixtures; no user workbook is opened or edited."""
    folder = tmp_path_factory.mktemp("native-excel")
    paths = {}
    try:
        with com_excel.excel_application() as application:
            for extension, file_format in (
                (".xlsx", 51),
                (".xlsm", 52),
                (".xls", 56),
                (".xlsb", 50),
            ):
                book = application.Workbooks.Add(-4167)  # xlWBATWorksheet
                try:
                    sheet = book.Worksheets.Item(1)
                    sheet.Name = "Data"
                    sheet.Range("B1:B4").Value2 = (("first",), (42,), (0,), (False,))
                    sheet.Range("B5").Formula = '=""'
                    sheet.Range("B6").Value2 = "must not be reached"
                    sheet.Range("C2").Formula = "=1+1"
                    sheet.Range("B1").Font.Name = "Arial"
                    sheet.Range("B1").Font.Size = 14
                    sheet.Range("B1").Font.Bold = True
                    sheet.Range("E1:F2").Merge()
                    book.Worksheets.Add().Name = "Other"
                    application.Calculate()
                    path = folder / f"users{extension}"
                    book.SaveAs(str(path), FileFormat=file_format)
                    paths[extension] = path
                finally:
                    sheet = None
                    book.Close(SaveChanges=False)
                    book = None
    except com_excel.ExcelComUnavailable as exc:
        pytest.skip(str(exc))
    return paths


@pytest.mark.parametrize("extension", [".xlsx", ".xlsm", ".xls", ".xlsb"])
def test_native_excel_reads_and_saves_original_format(
    native_workbooks, tmp_path, extension
):
    source = native_workbooks[extension]
    path = tmp_path / f"users{extension}"
    original = source.read_bytes()
    path.write_bytes(original)
    if extension == ".xls":
        assert original.startswith(b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1")
        assert not com_excel.is_ooxml_workbook(path)
    reader = ExcelReader()
    reader.load(path)
    assert reader.backend == "com"
    assert set(reader.sheet_names) == {"Data", "Other"}
    reader.select_sheet("Data")
    reader.set_start_cell("b1")
    assert reader.estimate_total_items() == 4
    assert reader.current_text == "first"
    writer = ExcelResultWriter()
    sig, backup = writer.write_error(
        path, "Data", "B1", "NOT_FOUND\x1b\x00", signature(path), backend=reader.backend
    )
    _, same_backup = writer.write_error(
        path, "Data", "B2", "MULTIPLE_MATCHES", sig, backend=reader.backend
    )
    assert backup == same_backup and Path(backup).read_bytes() == original
    assert Path(backup).suffix == path.suffix
    # Save/reopen must not invalidate the running input snapshot.
    for expected in ("42", "0", "False", None):
        reader.advance()
        assert reader.current_text == expected
    reader.close()
    with com_excel.excel_application() as application:
        try:
            with com_excel.open_workbook(application, path, read_only=True) as book:
                try:
                    sheet = book.Worksheets.Item("Data")
                    try:
                        _verify_native_result(sheet)
                        expected_format = {
                            ".xlsx": 51,
                            ".xlsm": 52,
                            ".xls": 56,
                            ".xlsb": 50,
                        }
                        assert book.FileFormat == expected_format[extension]
                    finally:
                        sheet = None
                finally:
                    book = None
        finally:
            application = None
    assert not list(tmp_path.glob(".cellrelay-*"))


def _verify_native_result(sheet):
    try:
        assert sheet.Range("G1").Value2.startswith(HEADER + "\n")
        assert "NOT_FOUND" in sheet.Range("G1").Value2
        assert "\x1b" not in sheet.Range("G1").Value2
        assert "MULTIPLE_MATCHES" in sheet.Range("G2").Value2
        assert sheet.Range("C2").Formula == "=1+1"
        assert sheet.Range("C2").Value2 == 2
        assert sheet.Range("B1").Value2 == "first"
        assert sheet.Range("B1").Font.Name == "Arial"
        assert sheet.Range("B1").Font.Size == 14
        assert sheet.Range("B1").Font.Bold
        assert sheet.Range("E1:F2").MergeCells
    finally:
        sheet = None


def test_native_locked_workbook_does_not_close_owner(native_workbooks, tmp_path):
    path = tmp_path / "locked.xlsx"
    original = native_workbooks[".xlsx"].read_bytes()
    path.write_bytes(original)
    with com_excel.excel_application() as owner:
        try:
            with com_excel.open_workbook(owner, path, read_only=False) as book:
                try:
                    locked_bytes = path.read_bytes()
                    with pytest.raises(PermissionError, match="읽기 전용"):
                        ExcelResultWriter().write_error(
                            path,
                            "Data",
                            "B1",
                            "NOT_FOUND",
                            signature(path),
                            backend="com",
                        )
                    assert book.Name == path.name
                    assert book.Saved
                    assert path.read_bytes() == locked_bytes
                    backups = list(tmp_path.glob("*.cellrelay-backup-*"))
                    assert len(backups) == 1 and backups[0].read_bytes() == locked_bytes
                finally:
                    book = None
        finally:
            owner = None
    assert path.read_bytes() == original
    assert not list(tmp_path.glob(".cellrelay-*"))


def test_native_excel_worker_initializes_com_on_its_qthread(native_workbooks, tmp_path):
    path = tmp_path / "users.xls"
    original = native_workbooks[".xls"].read_bytes()
    path.write_bytes(original)
    received = []

    class SaveThread(QThread):
        def run(self):
            worker = ExcelWorker()
            worker.finished.connect(
                lambda *result: received.append(result),
                Qt.ConnectionType.DirectConnection,
            )
            worker.write_error(
                7,
                {
                    "path": path,
                    "sheet": "Data",
                    "address": "B1",
                    "message": "NOT_FOUND",
                    "expected_signature": signature(path),
                    "backend": "com",
                },
            )

    thread = SaveThread()
    thread.start()
    thread.wait()
    assert len(received) == 1
    run_id, success, sig, detail = received[0]
    assert run_id == 7 and success, detail
    assert sig == signature(path) and Path(detail).read_bytes() == original

"""Application controller joining Excel, browser-worker, persistence, and UI."""

from __future__ import annotations

import logging
from dataclasses import replace
from pathlib import Path

from PySide6.QtCore import QObject, Qt, QThread, Signal, Slot

from app.browser.browser_worker import BrowserWorker
from app.config.settings import AppSettings, RuntimeProgress, SettingsStore
from app.core.state import AppState, ProgressSnapshot, StateMachine
from app.excel.excel_reader import ExcelReader
from app.excel.result_writer import ExcelWorker, signature

logger = logging.getLogger(__name__)


class CellRelayController(QObject):
    """Coordinate modules while keeping the UI free of workflow logic."""

    state_changed = Signal(str)
    sheets_changed = Signal(list)
    progress_changed = Signal(object)
    message_changed = Signal(str)
    selector_test_result = Signal(bool, str)
    browser_status_changed = Signal(bool, str)
    target_ready_changed = Signal(bool)
    training_confirmed = Signal(str)
    next_start_cell_changed = Signal(str)
    shutdown_completed = Signal()

    _open_browser_command = Signal(str, str)
    _test_selector_command = Signal(str)
    _test_aws_page_command = Signal(str)
    _input_command = Signal(int, str, str, str)
    _wait_clear_command = Signal(int, str, int, int)
    _aws_assign_command = Signal(int, str, str)
    _shutdown_command = Signal()
    _stop_barrier_command = Signal(int)
    _write_excel_command = Signal(int, object)

    def __init__(self, project_root: Path) -> None:
        super().__init__()
        self._state_machine = StateMachine()
        self._excel = ExcelReader()
        self._store = SettingsStore(project_root / "config")
        self._settings = self._store.load_settings()
        self._progress = ProgressSnapshot()
        self._browser_open = False
        self._browser_url = ""
        self._aws_page_ready = False
        self._target_check_in_progress = False
        self._active = False
        self._paused = False
        self._resume_state = AppState.WAITING_FOR_CLEAR
        self._run_id = 0
        self._last_completed_cell = ""
        self._last_skipped_cell = ""
        self._shutting_down = False
        self._shutdown_started = False
        self._confirmed_training_url = ""
        self._stopping = False
        self._stop_ack = False
        self._stop_error = ""
        self._excel_pending = False
        self._excel_write_kind = ""
        self._file_signature = None
        self._terminal_outcome = ""
        self._needs_advance = False
        self._last_error = ""
        self._error_report_backup = ""
        self._last_stage = ""

        self._excel_thread = QThread(self)
        self._excel_thread.setObjectName("CellRelayExcelThread")
        self._excel_worker = ExcelWorker()
        self._excel_worker.moveToThread(self._excel_thread)
        self._write_excel_command.connect(self._excel_worker.write_error)
        self._excel_worker.finished.connect(self._on_excel_written)
        self._excel_thread.finished.connect(self._excel_worker.deleteLater)
        self._excel_thread.finished.connect(self._on_shutdown_thread_finished)
        self._excel_thread.start()

        self._worker_thread = QThread(self)
        self._worker_thread.setObjectName("CellRelayBrowserThread")
        self._worker = BrowserWorker()
        self._worker.moveToThread(self._worker_thread)

        self._open_browser_command.connect(self._worker.open_browser)
        self._test_selector_command.connect(self._worker.test_selector)
        self._test_aws_page_command.connect(self._worker.test_aws_page)
        self._input_command.connect(self._worker.input_text)
        self._wait_clear_command.connect(self._worker.wait_for_clear)
        self._aws_assign_command.connect(self._worker.assign_aws_user)
        self._shutdown_command.connect(self._worker.shutdown)
        self._stop_barrier_command.connect(self._worker.acknowledge_stop)
        self._worker.stop_acknowledged.connect(self._on_stop_acknowledged)

        self._worker.browser_opened.connect(self._on_browser_opened)
        self._worker.selector_tested.connect(self._on_selector_tested)
        self._worker.aws_page_tested.connect(self._on_aws_page_tested)
        self._worker.text_inputted.connect(self._on_text_inputted)
        self._worker.clear_detected.connect(self._on_clear_detected)
        self._worker.assignment_stage_changed.connect(self._on_assignment_stage_changed)
        self._worker.assignment_completed.connect(self._on_assignment_completed)
        self._worker.assignment_skipped.connect(self._on_assignment_skipped)
        self._worker.operation_cancelled.connect(self._on_operation_cancelled)
        self._worker.operation_failed.connect(self._on_worker_failure)
        self._worker.shutdown_finished.connect(
            self._worker_thread.quit,
            Qt.ConnectionType.DirectConnection,
        )
        self._worker_thread.finished.connect(self._worker.deleteLater)
        self._worker_thread.finished.connect(self._on_shutdown_thread_finished)
        self._worker_thread.start()

    @property
    def settings(self) -> AppSettings:
        return replace(self._settings)

    @property
    def state(self) -> AppState:
        return self._state_machine.state

    @property
    def progress(self) -> ProgressSnapshot:
        return replace(self._progress)

    @property
    def aws_page_ready(self) -> bool:
        return self._aws_page_ready

    def load_excel(self, file_path: str) -> None:
        if self._active:
            self._set_message("작업을 중지한 뒤 Excel 파일을 변경하세요.")
            return
        try:
            self._transition(AppState.LOADING_EXCEL)
            self._set_message("Excel 파일을 불러오는 중입니다.")
            before = signature(Path(file_path))
            self._excel.load(file_path)
            self._file_signature = signature(self._excel.path)
            if before != self._file_signature:
                raise RuntimeError(
                    "불러오는 동안 Excel 파일이 변경되었습니다. 다시 시도하세요."
                )
            sheet_names = self._excel.sheet_names
            if not sheet_names:
                raise RuntimeError("Excel 파일에 Sheet가 없습니다.")
            preferred_sheet = self._settings.sheet
            if preferred_sheet not in sheet_names:
                preferred_sheet = sheet_names[0]
            self._excel.select_sheet(preferred_sheet)
            self._settings.excel_file = str(self._excel.path or file_path)
            self._settings.sheet = preferred_sheet
            self._save_settings()
            self.sheets_changed.emit(sheet_names)
            self._transition(AppState.READY)
            self._set_message(f"Excel 파일을 열었습니다: {Path(file_path).name}")
            logger.info("Excel loaded (%s): %s", self._excel.backend, file_path)
            logger.info("Sheet selected: %s", preferred_sheet)
        except Exception as exc:
            logger.exception("Excel load failed: %s", file_path)
            self._transition_to_error(f"Excel 파일을 열지 못했습니다: {exc}")

    def select_sheet(self, sheet_name: str) -> None:
        if not sheet_name or self._active:
            return
        try:
            self._excel.select_sheet(sheet_name)
            self._settings.sheet = sheet_name
            self._save_settings()
            self._set_message(f"Sheet를 선택했습니다: {sheet_name}")
            logger.info("Sheet selected: %s", sheet_name)
        except Exception as exc:
            logger.exception("Sheet selection failed: %s", sheet_name)
            self._transition_to_error(str(exc))

    def open_browser(self, url: str) -> None:
        url = url.strip()
        if not url:
            self._set_message("URL을 입력하세요.")
            return
        if self._active or self._target_check_in_progress:
            self._set_message("작업을 중지한 뒤 브라우저를 다시 여세요.")
            return
        self._settings.url = url
        self._save_settings()
        self._target_check_in_progress = False
        self._set_aws_page_ready(False)
        self._set_message("브라우저를 열고 페이지에 접속하는 중입니다.")
        self._open_browser_command.emit(url, self._settings.browser_channel)

    def test_target(
        self,
        workflow_mode: str,
        selector: str,
        url: str = "",
    ) -> None:
        if not self._browser_open:
            self._set_message("먼저 브라우저를 여세요.")
            self.selector_test_result.emit(False, "먼저 브라우저를 여세요.")
            return
        if self._active:
            self._set_message("작업을 중지한 뒤 페이지를 확인하세요.")
            return
        if workflow_mode == "aws_skill_builder":
            if self._target_check_in_progress:
                self._set_message("AWS 페이지 확인이 이미 진행 중입니다.")
                return
            self._set_aws_page_ready(False)
            self._target_check_in_progress = True
            self._worker.reset_control_flags()
            self._set_message(
                "수동 로그인과 교육 상세 페이지 준비 상태를 확인하는 중입니다."
            )
            # Discover the manually opened detail page, not the login URL.
            self._test_aws_page_command.emit("")
            return

        selector = selector.strip()
        if not selector:
            self._set_message("Text Selector를 입력하세요.")
            return
        self._settings.selector = selector
        self._save_settings()
        self._set_message("Text 영역을 확인하는 중입니다.")
        self._test_selector_command.emit(selector)

    def start_job(
        self,
        file_path: str,
        sheet: str,
        start_cell: str,
        url: str,
        selector: str,
        workflow_mode: str,
    ) -> None:
        if self._active or self._target_check_in_progress or self._shutting_down:
            self._set_message("이미 작업이 진행 중입니다.")
            return
        if not self._browser_open:
            self._set_message("먼저 브라우저를 열고 Text 영역을 확인하세요.")
            return
        try:
            requested_url = url.strip()
            if not requested_url:
                raise ValueError("URL을 입력하세요.")
            if requested_url != self._browser_url:
                raise ValueError(
                    "URL이 변경되었습니다. 브라우저 열기를 다시 실행하세요."
                )
            if workflow_mode not in {"text_clear", "aws_skill_builder"}:
                raise ValueError(f"지원하지 않는 동작 방식입니다: {workflow_mode}")
            if workflow_mode == "aws_skill_builder" and not self._aws_page_ready:
                raise ValueError(
                    "수동 로그인 후 교육 상세 페이지에서 "
                    "'AWS 페이지 확인'을 먼저 실행하세요."
                )
            requested_path = Path(file_path).expanduser().resolve()
            self.load_excel(str(requested_path))
            if self.state is AppState.ERROR:
                return
            self._excel.select_sheet(sheet)
            self._excel.set_start_cell(start_cell)
            total_items = self._excel.estimate_total_items()

            self._settings.excel_file = str(requested_path)
            self._settings.sheet = sheet
            self._settings.start_cell = self._excel.current_cell_address
            self._settings.url = requested_url
            self._settings.selector = selector.strip()
            self._settings.workflow_mode = workflow_mode
            if workflow_mode == "text_clear" and not self._settings.selector:
                raise ValueError("Text Selector를 입력하세요.")
            self._save_settings()

            if self.state is AppState.ERROR or self.state is AppState.COMPLETED:
                self._transition(AppState.READY)

            self._run_id += 1
            self._worker.reset_control_flags()
            self._active = True
            self._paused = False
            self._last_completed_cell = ""
            self._last_skipped_cell = ""
            self._terminal_outcome = ""
            self._needs_advance = False
            self._last_error = ""
            self._stop_error = ""
            self._error_report_backup = ""
            self._last_stage = ""
            self._progress = ProgressSnapshot(
                current_cell=self._excel.current_cell_address,
                current_value=self._excel.current_text or "",
                processed_count=0,
                total_items=total_items,
                last_message="작업을 시작합니다.",
            )
            self._emit_progress()
            self._save_runtime("READY_TO_INPUT")
            self._queue_current_value()
        except Exception as exc:
            self._active = False
            logger.exception("Could not start job")
            self._transition_to_error(f"작업을 시작하지 못했습니다: {exc}")

    def pause_job(self) -> None:
        if not self._active or self._paused:
            return
        if self.state not in {AppState.INPUTTING, AppState.WAITING_FOR_CLEAR}:
            return
        self._resume_state = self.state
        self._paused = True
        self._worker.request_pause()
        self._transition(AppState.PAUSED)
        self._set_message("작업을 일시정지했습니다.")
        self._save_runtime_safely("PAUSED")
        logger.info("Job paused at %s", self._progress.current_cell)

    def resume_job(self) -> None:
        if not self._active or not self._paused:
            return
        self._paused = False
        self._worker.request_resume()
        self._transition(self._resume_state)
        if self._resume_state is AppState.WAITING_FOR_CLEAR:
            self._set_message("Text 영역이 비워지기를 기다리는 중입니다.")
        else:
            self._set_message("Text 입력을 재개했습니다.")
        self._save_runtime_safely(self._resume_state.value)
        logger.info("Job resumed at %s", self._progress.current_cell)
        if self._needs_advance:
            self._advance_after_receipt()

    def stop_job(self) -> None:
        if not self._active or self._stopping:
            return
        self._begin_stop()

    def _begin_stop(self, error: str = "") -> None:
        if error:
            self._stop_error = error
            self._last_error = error
        if self._stopping:
            return
        self._stopping = True
        self._stop_ack = False
        self._paused = False
        self._worker.request_stop()
        self._transition(AppState.STOPPING)
        self._set_message("중지 요청을 처리하고 진행 기록을 저장하는 중입니다.")
        self._save_runtime_safely("STOPPING")
        self._stop_barrier_command.emit(self._run_id)

    @Slot(int)
    def _on_stop_acknowledged(self, run_id: int) -> None:
        if run_id == self._run_id and self._stopping:
            self._stop_ack = True
            self._finish_stop_if_ready()

    def _finish_stop_if_ready(self) -> None:
        if not self._stopping or not self._stop_ack or self._excel_pending:
            return
        if self._needs_advance:
            self._advance_after_receipt(continue_run=False)
        self._active = False
        self._stopping = False
        self._run_id += 1
        if (
            not self._terminal_outcome
            and self._last_stage
            in {"ASSIGNING_USER", "CONFIRMING_ASSIGNMENT", "ASSIGNMENT_SUBMITTED"}
            and not self._stop_error
        ):
            self._stop_error = "할당 도중 중지했습니다. 웹에서 실제 등록 여부를 확인한 뒤 재시작 셀을 결정하세요."
            self._last_error = "STOPPED_UNCONFIRMED: " + self._stop_error
        self._transition(AppState.ERROR if self._stop_error else AppState.READY)
        self._set_message(
            self._stop_error
            or "작업을 중지했습니다. 확인/기록 완료된 행만 이동했습니다."
        )
        self._save_runtime_safely("ERROR" if self._stop_error else "STOPPED")
        self.next_start_cell_changed.emit(self._progress.current_cell)
        if self._shutting_down:
            self._launch_shutdown()

    def shutdown(self) -> None:
        if self._shutting_down:
            return
        self._shutting_down = True
        if self._active:
            self.stop_job()
        else:
            self._launch_shutdown()

    def _launch_shutdown(self) -> None:
        if self._shutdown_started:
            return
        self._shutdown_started = True
        self._worker.request_stop()
        self._excel.close()
        self._shutdown_command.emit()
        self._excel_thread.quit()

    @Slot()
    def _on_shutdown_thread_finished(self) -> None:
        if (
            self._shutting_down
            and not self._worker_thread.isRunning()
            and not self._excel_thread.isRunning()
        ):
            self.shutdown_completed.emit()

    def wait_for_shutdown(self, timeout_ms: int = 5_000) -> bool:
        return self._worker_thread.wait(timeout_ms) and self._excel_thread.wait(
            timeout_ms
        )

    @Slot(str, str)
    def _on_browser_opened(self, url: str, browser_name: str) -> None:
        self._browser_open = True
        self._browser_url = url
        if self.state is AppState.ERROR:
            target = AppState.READY if self._excel.path is not None else AppState.IDLE
            self._transition(target)
        message = (
            f"브라우저를 열었습니다. ({browser_name}) "
            "필요하면 수동 로그인 후 교육 상세 페이지에서 AWS 페이지 확인을 누르세요."
        )
        self._set_message(message)
        self.browser_status_changed.emit(True, message)
        logger.info("Browser ready: %s", url)

    @Slot(bool, str, int)
    def _on_selector_tested(self, success: bool, message: str, count: int) -> None:
        self._set_message(message)
        self.selector_test_result.emit(success, message)
        if success:
            logger.info("Selector test passed with %s element", count)
        else:
            logger.warning("Selector test failed: %s", message)

    @Slot(bool, str, str)
    def _on_aws_page_tested(
        self,
        success: bool,
        message: str,
        actual_url: str,
    ) -> None:
        self._target_check_in_progress = False
        self._set_aws_page_ready(success)
        if success:
            self._confirmed_training_url = actual_url
            self.training_confirmed.emit(actual_url)
        self._set_message(message)
        self.selector_test_result.emit(success, message)
        if success:
            logger.info("AWS page test passed: %s", actual_url)
        else:
            logger.warning("AWS page test failed: %s", message)

    @Slot(int, str)
    def _on_text_inputted(self, run_id: int, actual_value: str) -> None:
        if run_id != self._run_id or not self._active:
            return
        if self._stopping:
            self._save_runtime_safely("INPUTTED_BEFORE_STOP")
            return
        logger.info("%s inserted", self._excel.current_cell_address)
        self._progress.current_value = actual_value
        self._resume_state = AppState.WAITING_FOR_CLEAR
        if not self._paused:
            self._transition(AppState.WAITING_FOR_CLEAR)
            self._set_message("Text 영역이 비워지기를 기다리는 중입니다.")
        else:
            self._set_message("입력 완료 후 일시정지 상태입니다.")
        try:
            self._save_runtime("WAITING_FOR_CLEAR")
        except OSError as exc:
            self._fail_active_job(f"진행 상태를 저장하지 못했습니다: {exc}")
            return
        self._wait_clear_command.emit(
            run_id,
            self._settings.selector,
            self._settings.stable_empty_ms,
            self._settings.status_check_interval_ms,
        )

    @Slot(int)
    def _on_clear_detected(self, run_id: int) -> None:
        self._complete_current_item(run_id, "Text 영역 clear 확인")

    @Slot(int, str, str)
    def _on_assignment_stage_changed(
        self,
        run_id: int,
        stage: str,
        message: str,
    ) -> None:
        if run_id != self._run_id or not self._active:
            return
        self._last_stage = stage
        self._set_message(message)
        self._save_runtime_safely(stage)

    @Slot(int)
    def _on_assignment_completed(self, run_id: int) -> None:
        self._complete_current_item(
            run_id, "AWS 완료 버튼 처리 완료 (관리자 명단 대조 필요)"
        )

    @Slot(int, str)
    def _on_assignment_skipped(self, run_id: int, reason: str) -> None:
        if run_id != self._run_id or not self._active:
            return
        self._record_excel_error(reason, "skip")

    def _record_excel_error(self, reason: str, kind: str) -> None:
        """A terminal skip is acknowledged only after the background save."""
        if self._excel_pending or self._terminal_outcome:
            return
        try:
            self._last_error = reason
            self._save_runtime(
                "AWS_ASSIGNMENT_SKIPPED_PENDING_EXCEL"
                if kind == "skip"
                else "ERROR_PENDING_EXCEL"
            )
            self._excel_pending = True
            self._excel_write_kind = kind
            self._set_message(
                "Excel의 별도 오류 열에 원인과 처리 시간을 저장하는 중입니다."
            )
            self._write_excel_command.emit(
                self._run_id,
                {
                    "path": self._excel.path,
                    "sheet": self._settings.sheet,
                    "address": self._excel.current_cell_address,
                    "message": reason,
                    "expected_signature": self._file_signature,
                    "backend": self._excel.backend,
                },
            )
        except Exception as exc:
            logger.exception("Could not queue Excel report")
            self._fail_active_job(f"Excel 오류 기록을 시작하지 못했습니다: {exc}")

    @Slot(int, bool, object, str)
    def _on_excel_written(self, run_id: int, success: bool, sig, detail: str) -> None:
        if run_id != self._run_id or not self._excel_pending:
            return
        self._excel_pending = False
        if not success:
            self._fail_active_job(
                f"Excel 오류 열 저장 실패: {detail}. 현재 셀에서 중지합니다."
            )
        else:
            self._file_signature = sig
            self._error_report_backup = detail
            if self._excel_write_kind == "skip":
                self._complete_current_item(
                    run_id, "AWS_ASSIGNMENT_SKIPPED", skipped=True
                )
            else:
                self._save_runtime_safely("ERROR_RECORDED")
        self._finish_stop_if_ready()

    def _complete_current_item(
        self, run_id: int, reason: str, skipped: bool = False
    ) -> None:
        if run_id != self._run_id or not self._active:
            return
        if self._terminal_outcome:
            return
        try:
            completed_cell = self._excel.current_cell_address
            if skipped:
                self._last_skipped_cell = completed_cell
                self._progress.skipped_count += 1
            else:
                self._last_completed_cell = completed_cell
            self._progress.processed_count += 1
            logger.info("Current item completed (%s): %s", reason, completed_cell)

            # Persist completion before advancing. If the process stops here,
            # recovery data tells future code not to resend this cell.
            completion_phase = (
                "AWS_ASSIGNMENT_SKIPPED"
                if skipped
                else (
                    "AWS_ASSIGNMENT_SUBMITTED"
                    if self._settings.workflow_mode == "aws_skill_builder"
                    else "CLEARED"
                )
            )
            self._terminal_outcome = completion_phase
            self._save_runtime(completion_phase)
            self._needs_advance = True
            self._emit_progress()
            if not self._paused and not self._stopping:
                self._advance_after_receipt()
        except Exception as exc:
            logger.exception("Could not save terminal receipt")
            # Keep the terminal outcome in memory; do not issue another action.
            self._fail_active_job(
                f"처리 결과 저장 실패: {exc}. 로그와 웹 결과를 확인하세요."
            )

    def _advance_after_receipt(self, continue_run: bool = True) -> None:
        try:
            if not self._needs_advance:
                return
            next_cell = self._excel.advance()
            self._needs_advance = False
            logger.info("Moving to %s", next_cell)
            self._progress.current_cell = next_cell
            self._progress.current_value = self._excel.current_text or ""
            self._emit_progress()
            self._save_runtime("READY_TO_INPUT")
            if continue_run:
                self._queue_current_value()
        except Exception as exc:
            logger.exception("Could not advance after clear")
            self._fail_active_job(f"다음 셀로 이동하지 못했습니다: {exc}")

    @Slot(int, str)
    def _on_operation_cancelled(self, run_id: int, operation: str) -> None:
        if run_id == self._run_id and self._active:
            logger.info("Browser operation cancelled: %s", operation)

    @Slot(str, int, str, str)
    def _on_worker_failure(
        self,
        operation: str,
        run_id: int,
        message: str,
        detail: str,
    ) -> None:
        logger.error("Browser failure (%s): %s\n%s", operation, message, detail)
        if operation == "open_browser":
            self._browser_open = False
            self._browser_url = ""
            self.browser_status_changed.emit(False, message)
            self._transition_to_error(f"브라우저를 열지 못했습니다: {message}")
            return
        if operation in {"input_text", "wait_for_clear", "aws_assign_user"}:
            if run_id != self._run_id or not self._active:
                return
            if operation == "aws_assign_user":
                self._record_excel_error(message, "error")
            self._fail_active_job(f"브라우저 작업에 실패했습니다: {message}")
            return
        self._set_message(message)

    def _queue_current_value(self) -> None:
        if self._stopping or self._paused or self._excel_pending:
            return
        self._terminal_outcome = ""
        self._last_stage = ""
        value = self._excel.current_text
        self._progress.current_cell = self._excel.current_cell_address
        self._progress.current_value = value or ""
        self._emit_progress()
        if value is None:
            self._active = False
            self._transition(AppState.COMPLETED)
            message = "빈 셀에 도달하여 작업을 완료했습니다."
            if self._settings.workflow_mode == "aws_skill_builder":
                message += " 관리자가 최종 등록 명단을 별도로 확보하여 대조하세요."
            self._set_message(message)
            self._save_runtime_safely("COMPLETED")
            logger.info(
                "Job completed before empty cell: %s", self._excel.current_cell_address
            )
            return
        self._transition(AppState.INPUTTING)
        if self._settings.workflow_mode == "aws_skill_builder":
            self._set_message(
                f"{self._excel.current_cell_address} 사용자 할당을 시작합니다."
            )
            self._save_runtime("OPENING_ASSIGNMENT")
            self._aws_assign_command.emit(
                self._run_id,
                value,
                self._confirmed_training_url,
            )
        else:
            self._set_message(
                f"{self._excel.current_cell_address} 값을 입력하는 중입니다."
            )
            self._input_command.emit(
                self._run_id,
                self._settings.selector,
                value,
                self._settings.input_method,
            )

    def _save_settings(self) -> None:
        self._store.save_settings(self._settings)

    def _save_runtime(self, phase: str) -> None:
        runtime = RuntimeProgress(
            excel_file=self._settings.excel_file,
            sheet=self._settings.sheet,
            start_cell=self._settings.start_cell,
            current_cell=self._progress.current_cell,
            current_value=self._progress.current_value,
            last_completed_cell=self._last_completed_cell,
            last_skipped_cell=self._last_skipped_cell,
            processed_count=self._progress.processed_count,
            skipped_count=self._progress.skipped_count,
            total_items=self._progress.total_items,
            url=self._settings.url,
            selector=self._settings.selector,
            workflow_mode=self._settings.workflow_mode,
            phase=phase,
            confirmed_training_url=self._confirmed_training_url,
            terminal_outcome=self._terminal_outcome,
            last_error=self._last_error,
            error_report_backup=self._error_report_backup,
        )
        self._store.save_progress(runtime)

    def _save_runtime_safely(self, phase: str) -> None:
        try:
            self._save_runtime(phase)
        except Exception:
            logger.exception("Could not save runtime progress: %s", phase)

    def _fail_active_job(self, message: str) -> None:
        if self._active:
            self._begin_stop(message)
        else:
            self._transition_to_error(message)

    def _transition_to_error(self, message: str) -> None:
        if self.state is not AppState.ERROR:
            self._transition(AppState.ERROR)
        self._set_message(message)

    def _transition(self, new_state: AppState) -> None:
        if self._state_machine.transition_to(new_state):
            self.state_changed.emit(new_state.value)
            logger.info("State changed: %s", new_state.value)

    def _set_message(self, message: str) -> None:
        self._progress.last_message = message
        self.message_changed.emit(message)
        self._emit_progress()

    def _set_aws_page_ready(self, ready: bool) -> None:
        if not ready:
            self._confirmed_training_url = ""
            self.training_confirmed.emit("")
        if self._aws_page_ready == ready:
            return
        self._aws_page_ready = ready
        self.target_ready_changed.emit(ready)

    def _emit_progress(self) -> None:
        self.progress_changed.emit(replace(self._progress))

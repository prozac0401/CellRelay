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

    _open_browser_command = Signal(str, str)
    _test_selector_command = Signal(str)
    _test_aws_page_command = Signal(str)
    _input_command = Signal(int, str, str, str)
    _wait_clear_command = Signal(int, str, int, int)
    _aws_assign_command = Signal(int, str, str)
    _shutdown_command = Signal()

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
        self._pending_aws_url = ""
        self._target_check_in_progress = False
        self._active = False
        self._paused = False
        self._resume_state = AppState.WAITING_FOR_CLEAR
        self._run_id = 0
        self._last_completed_cell = ""
        self._last_skipped_cell = ""
        self._clear_pending = False
        self._assignment_skip_pending_reason: str | None = None
        self._shutting_down = False

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
            self._excel.load(file_path)
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
            logger.info("Excel loaded: %s", file_path)
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
        if self._active:
            self._set_message("작업을 중지한 뒤 브라우저를 다시 여세요.")
            return
        self._settings.url = url
        self._save_settings()
        self._pending_aws_url = ""
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
            expected_url = url.strip()
            if not expected_url:
                self._set_message("AWS 교육 상세 URL을 입력하세요.")
                return
            self._set_aws_page_ready(False)
            self._pending_aws_url = expected_url
            self._target_check_in_progress = True
            self._worker.reset_control_flags()
            self._set_message(
                "수동 로그인과 교육 상세 페이지 준비 상태를 확인하는 중입니다."
            )
            self._test_aws_page_command.emit(expected_url)
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
        if self._active:
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
            if self._excel.path != requested_path:
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
            self._clear_pending = False
            self._assignment_skip_pending_reason = None
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
        if self._clear_pending:
            self._clear_pending = False
            self._on_clear_detected(self._run_id)
        elif self._assignment_skip_pending_reason is not None:
            reason = self._assignment_skip_pending_reason
            self._assignment_skip_pending_reason = None
            self._on_assignment_skipped(self._run_id, reason)

    def stop_job(self) -> None:
        if not self._active:
            return
        self._active = False
        self._paused = False
        self._clear_pending = False
        self._assignment_skip_pending_reason = None
        self._run_id += 1  # Ignore all already-queued results from the old run.
        self._worker.request_stop()
        self._transition(AppState.READY)
        self._set_message("작업을 중지했습니다. 현재 셀은 이동하지 않았습니다.")
        self._save_runtime_safely("STOPPED")
        logger.info("Job stopped at %s", self._progress.current_cell)

    def shutdown(self) -> None:
        if self._shutting_down:
            return
        self._shutting_down = True
        self._active = False
        self._worker.request_stop()
        self._excel.close()
        self._shutdown_command.emit()

    def wait_for_shutdown(self, timeout_ms: int = 5_000) -> bool:
        return self._worker_thread.wait(timeout_ms)

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
            self._browser_url = self._pending_aws_url
            self._settings.url = self._pending_aws_url
            self._save_settings()
        self._pending_aws_url = ""
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
        self._set_message(message)
        self._save_runtime_safely(stage)

    @Slot(int)
    def _on_assignment_completed(self, run_id: int) -> None:
        self._complete_current_item(run_id, "AWS 사용자 할당 확인")

    @Slot(int, str)
    def _on_assignment_skipped(self, run_id: int, reason: str) -> None:
        if run_id != self._run_id or not self._active:
            return
        if self._paused:
            self._assignment_skip_pending_reason = reason
            return
        try:
            skipped_cell = self._excel.current_cell_address
            self._save_runtime("AWS_ASSIGNMENT_SKIPPED_PENDING_EXCEL")
            self._excel.mark_current_cell_font_red()
            self._last_skipped_cell = skipped_cell
            self._progress.processed_count += 1
            self._progress.skipped_count += 1
            logger.warning(
                "AWS user was not assigned; marked %s red: %s",
                skipped_cell,
                reason,
            )

            # A skipped row is terminal only after the workbook records it.
            # Persist that fact before moving to the next Excel coordinate.
            self._save_runtime("AWS_ASSIGNMENT_SKIPPED")
            next_cell = self._excel.advance()
            logger.info("Moving to %s after skipped assignment", next_cell)
            self._progress.current_cell = next_cell
            self._progress.current_value = self._excel.current_text or ""
            self._emit_progress()
            self._save_runtime("READY_TO_INPUT")
            self._queue_current_value()
        except Exception as exc:
            logger.exception("Could not record and advance a skipped AWS user")
            self._fail_active_job(
                f"추가하지 못한 사용자를 Excel에 표시하지 못했습니다: {exc}"
            )

    def _complete_current_item(self, run_id: int, reason: str) -> None:
        if run_id != self._run_id or not self._active:
            return
        if self._paused:
            # Normally the worker gates this signal while paused. This check also
            # covers a signal that was already queued when Pause was clicked.
            self._clear_pending = True
            return
        try:
            completed_cell = self._excel.current_cell_address
            self._last_completed_cell = completed_cell
            self._progress.processed_count += 1
            logger.info("Current item completed (%s): %s", reason, completed_cell)

            # Persist completion before advancing. If the process stops here,
            # recovery data tells future code not to resend this cell.
            completion_phase = (
                "AWS_ASSIGNMENT_CONFIRMED"
                if self._settings.workflow_mode == "aws_skill_builder"
                else "CLEARED"
            )
            self._save_runtime(completion_phase)
            next_cell = self._excel.advance()
            logger.info("Moving to %s", next_cell)
            self._progress.current_cell = next_cell
            self._progress.current_value = self._excel.current_text or ""
            self._emit_progress()
            self._save_runtime("READY_TO_INPUT")
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
            self._fail_active_job(f"브라우저 작업에 실패했습니다: {message}")
            return
        self._set_message(message)

    def _queue_current_value(self) -> None:
        value = self._excel.current_text
        self._progress.current_cell = self._excel.current_cell_address
        self._progress.current_value = value or ""
        self._emit_progress()
        if value is None:
            self._active = False
            self._transition(AppState.COMPLETED)
            self._set_message("빈 셀에 도달하여 작업을 완료했습니다.")
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
            self._save_runtime_safely("OPENING_ASSIGNMENT")
            self._aws_assign_command.emit(
                self._run_id,
                value,
                self._settings.url,
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
        )
        self._store.save_progress(runtime)

    def _save_runtime_safely(self, phase: str) -> None:
        try:
            self._save_runtime(phase)
        except Exception:
            logger.exception("Could not save runtime progress: %s", phase)

    def _fail_active_job(self, message: str) -> None:
        self._active = False
        self._paused = False
        self._worker.request_stop()
        self._save_runtime_safely("ERROR")
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
        if self._aws_page_ready == ready:
            return
        self._aws_page_ready = ready
        self.target_ready_changed.emit(ready)

    def _emit_progress(self) -> None:
        self.progress_changed.emit(replace(self._progress))

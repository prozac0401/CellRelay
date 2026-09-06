"""Main CellRelay window; workflow decisions remain in the controller."""

from __future__ import annotations

from pathlib import Path

from PySide6.QtCore import QTimer
from PySide6.QtGui import QCloseEvent
from PySide6.QtWidgets import (
    QComboBox,
    QFileDialog,
    QFormLayout,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QSizePolicy,
    QVBoxLayout,
    QWidget,
)

from app import __version__
from app.core.controller import CellRelayController
from app.core.state import AppState, ProgressSnapshot


class MainWindow(QMainWindow):
    """Render state and forward user intent to CellRelayController."""

    def __init__(self, controller: CellRelayController) -> None:
        super().__init__()
        self._controller = controller
        self._saved_settings = controller.settings
        self._aws_page_ready = controller.aws_page_ready
        self._closing = False
        self._can_close = False
        self.setWindowTitle(f"CellRelay {__version__}")
        self.setMinimumSize(720, 610)
        self.resize(780, 660)

        self._build_ui()
        self._connect_signals()
        self._apply_settings()
        self._on_state_changed(controller.state.value)
        self._on_progress_changed(controller.progress)

        saved_file = Path(self._saved_settings.excel_file)
        if saved_file.is_file() and saved_file.suffix.lower() == ".xlsx":
            QTimer.singleShot(0, lambda: self._controller.load_excel(str(saved_file)))

    def _build_ui(self) -> None:
        central = QWidget(self)
        root_layout = QVBoxLayout(central)
        root_layout.setContentsMargins(18, 18, 18, 18)
        root_layout.setSpacing(14)

        excel_group = QGroupBox("Excel")
        excel_layout = QGridLayout(excel_group)
        self.excel_path_edit = QLineEdit()
        self.excel_path_edit.setReadOnly(True)
        self.browse_button = QPushButton("찾기")
        self.sheet_combo = QComboBox()
        self.start_cell_edit = QLineEdit("B5")
        self.start_cell_edit.setMaximumWidth(140)
        excel_layout.addWidget(QLabel("Excel 파일"), 0, 0)
        excel_layout.addWidget(self.excel_path_edit, 0, 1)
        excel_layout.addWidget(self.browse_button, 0, 2)
        excel_layout.addWidget(QLabel("Sheet"), 1, 0)
        excel_layout.addWidget(self.sheet_combo, 1, 1, 1, 2)
        excel_layout.addWidget(QLabel("시작 셀"), 2, 0)
        excel_layout.addWidget(self.start_cell_edit, 2, 1, 1, 2)

        browser_group = QGroupBox("Browser")
        browser_layout = QFormLayout(browser_group)
        self.url_edit = QLineEdit()
        self.url_edit.setPlaceholderText("https://example.com")
        self.selector_edit = QLineEdit("textarea")
        self.selector_edit.setPlaceholderText("textarea 또는 #prompt-textarea")
        self.workflow_combo = QComboBox()
        self.workflow_combo.addItem(
            "AWS Skill Builder 사용자 자동 할당",
            "aws_skill_builder",
        )
        self.workflow_combo.addItem(
            "Text Clear 감지 (수동 Action)",
            "text_clear",
        )
        browser_layout.addRow("URL", self.url_edit)
        browser_layout.addRow("동작 방식", self.workflow_combo)
        browser_layout.addRow("Text Selector", self.selector_edit)
        self.training_url_edit = QLineEdit()
        self.training_url_edit.setReadOnly(True)
        self.training_url_edit.setPlaceholderText(
            "수동 로그인 → 과정 상세페이지 이동 → AWS 페이지 확인"
        )
        browser_layout.addRow("확인된 과정", self.training_url_edit)

        browser_buttons = QHBoxLayout()
        self.open_browser_button = QPushButton("브라우저 열기")
        self.test_selector_button = QPushButton("Text 영역 확인")
        browser_buttons.addWidget(self.open_browser_button)
        browser_buttons.addWidget(self.test_selector_button)
        browser_buttons.addStretch(1)

        job_buttons = QHBoxLayout()
        self.start_button = QPushButton("시작")
        self.pause_button = QPushButton("일시정지")
        self.resume_button = QPushButton("재개")
        self.stop_button = QPushButton("중지")
        for button in (
            self.start_button,
            self.pause_button,
            self.resume_button,
            self.stop_button,
        ):
            button.setMinimumHeight(36)
            button.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
            job_buttons.addWidget(button)

        status_group = QGroupBox("실행 상태")
        status_layout = QGridLayout(status_group)
        self.state_value_label = QLabel(AppState.IDLE.value)
        self.cell_value_label = QLabel("-")
        self.current_value_label = QLabel("")
        self.current_value_label.setWordWrap(True)
        self.current_value_label.setTextInteractionFlags(
            self.current_value_label.textInteractionFlags()
        )
        self.count_value_label = QLabel("0")
        self.skipped_value_label = QLabel("0")
        self.message_value_label = QLabel("대기 중입니다.")
        self.message_value_label.setWordWrap(True)
        status_layout.addWidget(QLabel("현재 상태:"), 0, 0)
        status_layout.addWidget(self.state_value_label, 0, 1)
        status_layout.addWidget(QLabel("현재 셀:"), 1, 0)
        status_layout.addWidget(self.cell_value_label, 1, 1)
        status_layout.addWidget(QLabel("현재 값:"), 2, 0)
        status_layout.addWidget(self.current_value_label, 2, 1)
        status_layout.addWidget(QLabel("처리 개수:"), 3, 0)
        status_layout.addWidget(self.count_value_label, 3, 1)
        status_layout.addWidget(QLabel("추가 안 됨:"), 4, 0)
        status_layout.addWidget(self.skipped_value_label, 4, 1)
        status_layout.addWidget(QLabel("마지막 메시지:"), 5, 0)
        status_layout.addWidget(self.message_value_label, 5, 1)
        status_layout.setColumnStretch(1, 1)

        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 1)
        self.progress_bar.setValue(0)
        self.progress_bar.setFormat("0 / ?")

        root_layout.addWidget(excel_group)
        root_layout.addWidget(browser_group)
        root_layout.addLayout(browser_buttons)
        root_layout.addLayout(job_buttons)
        root_layout.addWidget(status_group)
        root_layout.addWidget(self.progress_bar)
        root_layout.addStretch(1)
        self.setCentralWidget(central)

    def _connect_signals(self) -> None:
        self.browse_button.clicked.connect(self._choose_excel_file)
        self.sheet_combo.currentTextChanged.connect(self._controller.select_sheet)
        self.open_browser_button.clicked.connect(self._open_browser)
        self.test_selector_button.clicked.connect(self._test_target)
        self.start_button.clicked.connect(self._start_job)
        self.pause_button.clicked.connect(self._controller.pause_job)
        self.resume_button.clicked.connect(self._controller.resume_job)
        self.stop_button.clicked.connect(self._controller.stop_job)
        self.workflow_combo.currentIndexChanged.connect(self._update_mode_controls)

        self._controller.state_changed.connect(self._on_state_changed)
        self._controller.sheets_changed.connect(self._on_sheets_changed)
        self._controller.progress_changed.connect(self._on_progress_changed)
        self._controller.message_changed.connect(self.message_value_label.setText)
        self._controller.target_ready_changed.connect(self._on_target_ready_changed)
        self._controller.training_confirmed.connect(self.training_url_edit.setText)
        self._controller.next_start_cell_changed.connect(self.start_cell_edit.setText)
        self._controller.shutdown_completed.connect(self._finish_close)

    def _apply_settings(self) -> None:
        settings = self._saved_settings
        self.excel_path_edit.setText(settings.excel_file)
        self.start_cell_edit.setText(settings.start_cell or "B5")
        self.url_edit.setText(settings.url)
        self.selector_edit.setText(settings.selector or "textarea")
        mode_index = self.workflow_combo.findData(settings.workflow_mode)
        self.workflow_combo.setCurrentIndex(max(0, mode_index))
        self._update_mode_controls()

    def _choose_excel_file(self) -> None:
        initial_dir = self.excel_path_edit.text() or str(Path.home())
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Excel 파일 선택",
            initial_dir,
            "Excel Workbook (*.xlsx)",
        )
        if not file_path:
            return
        self.excel_path_edit.setText(file_path)
        self._controller.load_excel(file_path)

    def _open_browser(self) -> None:
        self._controller.open_browser(self.url_edit.text())

    def _test_target(self) -> None:
        self._controller.test_target(
            str(self.workflow_combo.currentData()),
            self.selector_edit.text(),
            self.url_edit.text(),
        )

    def _start_job(self) -> None:
        workflow_mode = str(self.workflow_combo.currentData())
        if workflow_mode == "aws_skill_builder":
            answer = QMessageBox.question(
                self,
                "AWS 사용자 자동 할당 시작",
                "Excel 시작 셀부터 첫 빈 셀까지 사용자를 이 교육에 자동 할당합니다.\n\n"
                "검색 결과가 정확히 1개이고 검색값과 일치할 때만 할당합니다. "
                "일치하지 않는 사용자는 취소하고 Excel의 'CellRelay 검색 오류' 열에 원인을 기록한 뒤 "
                "다음 행으로 진행합니다. "
                "검색 완료 또는 완료 버튼·팝업 종료를 확인하지 못하면 중지합니다.\n"
                "교육 상세 페이지의 사용자 표는 검사하지 않습니다. "
                "처리 후 관리자가 최종 등록 명단을 별도로 확보하여 대조하세요.\n\n"
                f"확인된 과정: {self.training_url_edit.text()}\n\n"
                "계속하시겠습니까?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No,
            )
            if answer != QMessageBox.StandardButton.Yes:
                return
        self._controller.start_job(
            file_path=self.excel_path_edit.text(),
            sheet=self.sheet_combo.currentText(),
            start_cell=self.start_cell_edit.text(),
            url=self.url_edit.text(),
            selector=self.selector_edit.text(),
            workflow_mode=workflow_mode,
        )

    def _on_sheets_changed(self, sheets: list[str]) -> None:
        self.sheet_combo.blockSignals(True)
        self.sheet_combo.clear()
        self.sheet_combo.addItems(sheets)
        selected = self._controller.settings.sheet
        if selected in sheets:
            self.sheet_combo.setCurrentText(selected)
        self.sheet_combo.blockSignals(False)

    def _on_state_changed(self, state: str) -> None:
        self.state_value_label.setText(state)
        active = state in {
            AppState.INPUTTING.value,
            AppState.WAITING_FOR_CLEAR.value,
            AppState.PAUSED.value,
            AppState.STOPPING.value,
        }
        self.browse_button.setEnabled(not active)
        self.sheet_combo.setEnabled(not active)
        self.start_cell_edit.setEnabled(not active)
        self.url_edit.setEnabled(not active)
        self.selector_edit.setEnabled(not active)
        self.workflow_combo.setEnabled(not active)
        self.open_browser_button.setEnabled(not active)
        self.test_selector_button.setEnabled(not active)
        self.start_button.setEnabled(
            state
            in {
                AppState.READY.value,
                AppState.COMPLETED.value,
                AppState.ERROR.value,
            }
        )
        self.pause_button.setEnabled(
            state in {AppState.INPUTTING.value, AppState.WAITING_FOR_CLEAR.value}
        )
        self.resume_button.setEnabled(state == AppState.PAUSED.value)
        self.stop_button.setEnabled(active and state != AppState.STOPPING.value)
        self._update_mode_controls()

    def _update_mode_controls(self) -> None:
        is_text_mode = self.workflow_combo.currentData() == "text_clear"
        active = self.state_value_label.text() in {
            AppState.INPUTTING.value,
            AppState.WAITING_FOR_CLEAR.value,
            AppState.PAUSED.value,
            AppState.STOPPING.value,
        }
        self.selector_edit.setEnabled(is_text_mode and not active)
        self.test_selector_button.setText(
            "Text 영역 확인" if is_text_mode else "AWS 페이지 확인"
        )
        self.test_selector_button.setEnabled(not active)
        state_allows_start = self.state_value_label.text() in {
            AppState.READY.value,
            AppState.COMPLETED.value,
            AppState.ERROR.value,
        }
        self.start_button.setEnabled(
            state_allows_start and (is_text_mode or self._aws_page_ready)
        )

    def _on_target_ready_changed(self, ready: bool) -> None:
        self._aws_page_ready = ready
        self._update_mode_controls()

    def _on_progress_changed(self, progress: ProgressSnapshot) -> None:
        self.cell_value_label.setText(progress.current_cell)
        self.current_value_label.setText(progress.current_value)
        self.skipped_value_label.setText(str(progress.skipped_count))
        if progress.total_items > 0:
            self.count_value_label.setText(
                f"{progress.processed_count} / {progress.total_items}"
            )
            self.progress_bar.setRange(0, progress.total_items)
            self.progress_bar.setValue(
                min(progress.processed_count, progress.total_items)
            )
            self.progress_bar.setFormat(
                f"{progress.processed_count} / {progress.total_items}"
            )
        else:
            self.count_value_label.setText(str(progress.processed_count))
            self.progress_bar.setRange(0, 1)
            self.progress_bar.setValue(0)
            self.progress_bar.setFormat(f"{progress.processed_count} / ?")
        self.message_value_label.setText(progress.last_message)

    def closeEvent(self, event: QCloseEvent) -> None:
        if self._can_close:
            event.accept()
            return
        event.ignore()
        if not self._closing:
            self._closing = True
            self.centralWidget().setEnabled(False)
            self.message_value_label.setText(
                "진행 기록 저장과 브라우저 종료를 기다리는 중입니다."
            )
            self._controller.shutdown()

    def _finish_close(self) -> None:
        self._can_close = True
        self.close()

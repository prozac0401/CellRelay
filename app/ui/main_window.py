"""Main CellRelay window; workflow decisions remain in the controller."""

from __future__ import annotations

from pathlib import Path

from PySide6.QtCore import QSize, Qt, QTimer
from PySide6.QtGui import QColor, QCloseEvent, QPaintEvent, QPainter, QPainterPath, QPen, QResizeEvent, QWheelEvent
from PySide6.QtWidgets import (
    QBoxLayout,
    QComboBox,
    QFileDialog,
    QFrame,
    QGridLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QScrollArea,
    QSizePolicy,
    QVBoxLayout,
    QWidget,
)

from app import __version__
from app.core.controller import CellRelayController
from app.core.state import AppState, ProgressSnapshot
from app.excel.com_excel import EXCEL_FILE_FILTER, SUPPORTED_EXTENSIONS
from app.ui.theme import apply_theme, set_style_property


class PlainTextLabel(QLabel):
    """Copyable untrusted text that never turns an Excel value into HTML.

    Ignoring the horizontal hint lets Qt wrap long paths, values and messages
    within the available width instead of expanding the whole application.
    """

    def __init__(self, text: str = "", *, selectable: bool = False) -> None:
        super().__init__(text)
        self.setTextFormat(Qt.TextFormat.PlainText)
        self.setWordWrap(True)
        self.setSizePolicy(QSizePolicy.Policy.Ignored, QSizePolicy.Policy.Preferred)
        if selectable:
            self.setTextInteractionFlags(
                Qt.TextInteractionFlag.TextSelectableByMouse
                | Qt.TextInteractionFlag.TextSelectableByKeyboard
            )
            self.setFocusPolicy(Qt.FocusPolicy.StrongFocus)

    def minimumSizeHint(self) -> QSize:
        size = super().minimumSizeHint()
        return QSize(0, size.height())


class ChoiceComboBox(QComboBox):
    """Keep the dropdown affordance visible on every native Qt platform.

    Qt stylesheet engines may omit the platform arrow when the drop-down
    border is styled. A tiny device-independent chevron avoids image assets
    while all popup, focus, keyboard and accessibility behavior stays native.
    """

    def paintEvent(self, event: QPaintEvent) -> None:
        super().paintEvent(event)
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        painter.setPen(QPen(QColor("#65738A" if self.isEnabled() else "#A1ADBF"), 1.5))
        center_x = self.width() - 16
        center_y = self.height() / 2
        path = QPainterPath()
        path.moveTo(center_x - 4, center_y - 2)
        path.lineTo(center_x, center_y + 2)
        path.lineTo(center_x + 4, center_y - 2)
        painter.drawPath(path)

    def wheelEvent(self, event: QWheelEvent) -> None:
        # A closed combo is part of the scrollable settings form: scrolling
        # must not silently select a different sheet or automation workflow.
        # Open popups retain Qt's native wheel behavior, as do keyboard keys.
        if self.view().isVisible():
            super().wheelEvent(event)
        else:
            event.ignore()


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
        self.setMinimumSize(560, 320)
        # QScreen reports logical pixels, so this also fits high-DPI laptop
        # work areas. Only the initial size is clamped; manual resizing stays
        # unrestricted above our compact minimum and uses the scrolling body.
        initial_size = QSize(1160, 840)
        screen = self.screen()
        if screen is not None:
            available = screen.availableGeometry()
            initial_size.setWidth(max(560, min(initial_size.width(), available.width() - 32)))
            initial_size.setHeight(max(320, min(initial_size.height(), available.height() - 48)))
        self.resize(initial_size)

        self._build_ui()
        apply_theme(self)
        self._connect_signals()
        self._apply_settings()
        self._on_state_changed(controller.state.value)
        self._on_progress_changed(controller.progress)

        saved_file = Path(self._saved_settings.excel_file)
        if saved_file.is_file() and saved_file.suffix.lower() in SUPPORTED_EXTENSIONS:
            QTimer.singleShot(0, lambda: self._controller.load_excel(str(saved_file)))

    def _build_ui(self) -> None:
        central = QWidget(self)
        central.setObjectName("AppRoot")
        root_layout = QVBoxLayout(central)
        root_layout.setContentsMargins(0, 0, 0, 0)
        root_layout.setSpacing(0)

        # Header and action bar stay outside the scrolling document. Settings
        # and status reflow as a single column on a compact desktop window.
        header = QWidget()
        header_layout = QHBoxLayout(header)
        header_layout.setContentsMargins(26, 22, 26, 18)
        header_layout.setSpacing(13)
        brand_mark = QLabel("CR")
        brand_mark.setObjectName("BrandMark")
        brand_mark.setAlignment(Qt.AlignmentFlag.AlignCenter)
        brand_mark.setFixedSize(46, 46)
        brand_layout = QVBoxLayout()
        brand_layout.setSpacing(2)
        brand_line = QHBoxLayout()
        brand_line.setSpacing(10)
        brand = QLabel("CellRelay")
        brand.setObjectName("Brand")
        version = QLabel(f"v{__version__}")
        version.setObjectName("Version")
        brand_line.addWidget(brand)
        brand_line.addWidget(version, 0, Qt.AlignmentFlag.AlignBottom)
        brand_line.addStretch()
        subtitle = PlainTextLabel("Excel에서 시작하는, 간편한 교육 할당")
        subtitle.setObjectName("Subtitle")
        brand_layout.addLayout(brand_line)
        brand_layout.addWidget(subtitle)
        header_layout.addWidget(brand_mark)
        header_layout.addLayout(brand_layout, 1)
        root_layout.addWidget(header)

        self.content_scroll = QScrollArea()
        self.content_scroll.setWidgetResizable(True)
        self.content_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.content_scroll.setFrameShape(QFrame.Shape.NoFrame)
        content = QWidget()
        content.setObjectName("ScrollContent")
        self._content_layout = QBoxLayout(QBoxLayout.Direction.LeftToRight, content)
        self._content_layout.setContentsMargins(26, 2, 26, 22)
        self._content_layout.setSpacing(20)
        self._content_layout.setAlignment(Qt.AlignmentFlag.AlignTop)
        self.settings_panel = QWidget()
        settings_layout = QVBoxLayout(self.settings_panel)
        settings_layout.setContentsMargins(0, 0, 0, 0)
        settings_layout.setSpacing(16)

        self.excel_group, excel_layout = self._section(
            "1", "Excel 데이터", "같은 열에서 아래로 읽으며, 첫 빈 셀에서 종료합니다."
        )
        self.excel_path_edit = QLineEdit()
        self.excel_path_edit.setReadOnly(True)
        self.excel_path_edit.setPlaceholderText("사용자 명단이 담긴 Excel 파일을 선택하세요")
        self.excel_path_edit.setToolTip("원본 Excel 파일 경로입니다. .xlsx, .xlsm, .xls, .xlsb 파일을 선택하세요.")
        self.browse_button = QPushButton("찾기")
        self.browse_button.setToolTip("Excel 통합 문서(.xlsx, .xlsm, .xls, .xlsb) 선택")
        self.sheet_combo = ChoiceComboBox()
        self.sheet_combo.setSizeAdjustPolicy(QComboBox.SizeAdjustPolicy.AdjustToMinimumContentsLengthWithIcon)
        self.sheet_combo.setMinimumContentsLength(8)
        self.sheet_combo.setToolTip("사용자 값을 읽을 Excel 시트를 선택하세요.")
        self.start_cell_edit = QLineEdit("B5")
        self.start_cell_edit.setPlaceholderText("예: B5")
        self.start_cell_edit.setToolTip("첫 번째 사용자 값의 셀 주소입니다. 예: B5")
        excel_layout.addWidget(self._field_label("Excel 파일", self.excel_path_edit))
        file_row = QHBoxLayout()
        file_row.setSpacing(8)
        file_row.addWidget(self.excel_path_edit, 1)
        file_row.addWidget(self.browse_button)
        excel_layout.addLayout(file_row)
        excel_fields = QGridLayout()
        excel_fields.setHorizontalSpacing(14)
        excel_fields.setVerticalSpacing(6)
        excel_fields.addWidget(self._field_label("Sheet", self.sheet_combo), 0, 0)
        excel_fields.addWidget(self._field_label("시작 셀", self.start_cell_edit), 0, 1)
        excel_fields.addWidget(self.sheet_combo, 1, 0)
        excel_fields.addWidget(self.start_cell_edit, 1, 1)
        excel_fields.setColumnStretch(0, 2)
        excel_fields.setColumnStretch(1, 1)
        excel_layout.addLayout(excel_fields)
        settings_layout.addWidget(self.excel_group)

        self.browser_group, browser_layout = self._section(
            "2", "브라우저 연결", "로그인과 과정 선택은 직접 진행해 주세요."
        )
        self.url_edit = QLineEdit()
        self.url_edit.setPlaceholderText("https://skillbuilder.aws/login")
        self.url_edit.setToolTip("브라우저 열기 버튼으로 접속할 URL입니다.")
        self.selector_edit = QLineEdit("textarea")
        self.selector_edit.setPlaceholderText("textarea 또는 #prompt-textarea")
        self.selector_edit.setToolTip("Text Clear 모드에서 사용할 입력 요소의 Playwright 선택자입니다.")
        self.workflow_combo = ChoiceComboBox()
        self.workflow_combo.setSizeAdjustPolicy(QComboBox.SizeAdjustPolicy.AdjustToMinimumContentsLengthWithIcon)
        self.workflow_combo.setMinimumContentsLength(12)
        self.workflow_combo.addItem(
            "AWS Skill Builder 사용자 자동 할당",
            "aws_skill_builder",
        )
        self.workflow_combo.addItem(
            "Text Clear 감지 (수동 Action)",
            "text_clear",
        )
        self.workflow_combo.setToolTip("AWS 자동 할당 또는 수동 Action을 사용하는 Text Clear 모드를 선택하세요.")
        browser_layout.addWidget(self._field_label("동작 방식", self.workflow_combo))
        browser_layout.addWidget(self.workflow_combo)
        browser_layout.addWidget(self._field_label("시작 URL", self.url_edit))
        browser_layout.addWidget(self.url_edit)
        self.selector_label = self._field_label("Text Selector", self.selector_edit)
        browser_layout.addWidget(self.selector_label)
        browser_layout.addWidget(self.selector_edit)
        self.training_url_edit = QLineEdit()
        self.training_url_edit.setReadOnly(True)
        self.training_url_edit.setPlaceholderText(
            "수동 로그인 → 과정 상세페이지 이동 → AWS 페이지 확인"
        )
        self.training_url_edit.setToolTip("AWS 페이지 확인으로 고정한 교육 상세 페이지입니다. 시작 전에 올바른 과정인지 확인하세요.")
        self.training_url_label = self._field_label("확인된 과정", self.training_url_edit)
        browser_layout.addWidget(self.training_url_label)
        browser_layout.addWidget(self.training_url_edit)

        browser_buttons = QHBoxLayout()
        browser_buttons.setSpacing(8)
        self.open_browser_button = QPushButton("브라우저 열기")
        self.test_selector_button = QPushButton("Text 영역 확인")
        self.open_browser_button.setToolTip("CellRelay가 제어하는 Edge 창을 엽니다. 로그인은 직접 진행하세요.")
        browser_buttons.addWidget(self.open_browser_button, 1)
        browser_buttons.addWidget(self.test_selector_button, 1)
        browser_layout.addSpacing(3)
        browser_layout.addLayout(browser_buttons)
        self.ready_indicator_label = PlainTextLabel("과정 상세페이지에 도착하면 AWS 페이지 확인을 눌러주세요.")
        self.ready_indicator_label.setObjectName("ReadyIndicator")
        browser_layout.addWidget(self.ready_indicator_label)
        settings_layout.addWidget(self.browser_group)
        settings_layout.addStretch(1)

        self.status_panel = QWidget()
        status_panel_layout = QVBoxLayout(self.status_panel)
        status_panel_layout.setContentsMargins(0, 0, 0, 0)
        status_panel_layout.setSpacing(16)
        self.status_group, status_layout = self._section("", "실행 상태", "")
        self.state_value_label = QLabel(AppState.IDLE.value)
        self.state_value_label.setObjectName("StateBadge")
        self.state_value_label.setTextFormat(Qt.TextFormat.PlainText)
        self.state_value_label.setAccessibleName("현재 상태")
        self.state_value_label.setSizePolicy(QSizePolicy.Policy.Maximum, QSizePolicy.Policy.Fixed)
        status_layout.addWidget(self.state_value_label)
        status_layout.addSpacing(9)

        self.cell_value_label = PlainTextLabel("-")
        self.cell_value_label.setObjectName("CellValue")
        self.count_value_label = PlainTextLabel("0")
        self.count_value_label.setObjectName("MetricValue")
        self.skipped_value_label = PlainTextLabel("0")
        self.skipped_value_label.setObjectName("MetricValue")
        metrics = QGridLayout()
        metrics.setHorizontalSpacing(10)
        metrics.setVerticalSpacing(2)
        for column, (title, value) in enumerate((
            ("현재 셀", self.cell_value_label),
            ("처리 개수", self.count_value_label),
            ("추가 안 됨", self.skipped_value_label),
        )):
            label = QLabel(title)
            label.setObjectName("MetricLabel")
            value.setTextFormat(Qt.TextFormat.PlainText)
            value.setAccessibleName(title)
            metrics.addWidget(label, 0, column)
            metrics.addWidget(value, 1, column)
            metrics.setColumnStretch(column, 1)
        status_layout.addLayout(metrics)

        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 1)
        self.progress_bar.setValue(0)
        self.progress_bar.setFormat("0 / ?")
        self.progress_bar.setTextVisible(False)
        self.progress_bar.setFixedHeight(8)
        self.progress_bar.setAccessibleName("전체 작업 진행률")
        status_layout.addSpacing(3)
        status_layout.addWidget(self.progress_bar)
        status_layout.addSpacing(10)
        self.current_value_label = PlainTextLabel("", selectable=True)
        self.current_value_label.setAccessibleName("현재 Excel 값")
        self.current_value_label.setToolTip("현재 처리 중인 Excel 셀 값입니다. 텍스트를 선택해 복사할 수 있습니다.")
        status_layout.addWidget(self._field_label("현재 값", self.current_value_label))
        self.current_value_scroll = self._text_view(self.current_value_label, 96)
        status_layout.addWidget(self.current_value_scroll)
        self.message_value_label = PlainTextLabel("대기 중입니다.", selectable=True)
        self.message_value_label.setAccessibleName("마지막 작업 메시지")
        status_layout.addWidget(self._field_label("마지막 메시지", self.message_value_label))
        self.message_value_scroll = self._text_view(self.message_value_label, 100)
        status_layout.addWidget(self.message_value_scroll)
        status_panel_layout.addWidget(self.status_group)

        self.guide_group, guide_layout = self._section("", "시작 전 확인", "")
        self.guide_label = PlainTextLabel(
            "1. Excel 파일 · Sheet · 시작 셀을 선택하세요.\n"
            "2. 브라우저에서 로그인 후 과정으로 이동하세요.\n"
            "3. AWS 페이지 확인 후 시작을 누르세요."
        )
        self.guide_label.setObjectName("Helper")
        guide_layout.addWidget(self.guide_label)
        self.audit_notice_label = PlainTextLabel(
            "처리 완료 후 최종 등록 명단은 관리자가 별도로 대조해야 합니다."
        )
        self.audit_notice_label.setObjectName("Helper")
        guide_layout.addSpacing(6)
        guide_layout.addWidget(self.audit_notice_label)
        status_panel_layout.addWidget(self.guide_group)
        status_panel_layout.addStretch(1)

        self._content_layout.addWidget(self.settings_panel, 3)
        self._content_layout.addWidget(self.status_panel, 2)
        self.content_scroll.setWidget(content)
        root_layout.addWidget(self.content_scroll, 1)

        self.action_bar = QFrame()
        self.action_bar.setObjectName("ActionBar")
        self._action_layout = QBoxLayout(QBoxLayout.Direction.LeftToRight, self.action_bar)
        self._action_layout.setContentsMargins(26, 16, 26, 16)
        self._action_layout.setSpacing(16)
        self.action_hint_label = PlainTextLabel("Excel 파일을 선택해 작업을 준비하세요.")
        self.action_hint_label.setObjectName("ActionHint")
        self._action_layout.addWidget(self.action_hint_label, 1)
        button_container = QWidget()
        job_buttons = QHBoxLayout(button_container)
        job_buttons.setContentsMargins(0, 0, 0, 0)
        job_buttons.setSpacing(8)
        self.start_button = QPushButton("시작")
        self.start_button.setObjectName("PrimaryButton")
        self.start_button.setToolTip("선택한 시작 셀부터 작업을 시작합니다. AWS 모드는 과정 확인이 필요합니다.")
        self.pause_button = QPushButton("일시정지")
        self.pause_button.setToolTip("현재 작업을 안전한 경계에서 잠시 멈춥니다.")
        self.resume_button = QPushButton("재개")
        self.resume_button.setToolTip("일시정지한 작업을 이어서 진행합니다.")
        self.stop_button = QPushButton("중지")
        self.stop_button.setObjectName("DangerButton")
        self.stop_button.setToolTip("진행 기록을 저장하고 현재 작업을 중지합니다.")
        for button in (self.start_button, self.pause_button, self.resume_button, self.stop_button):
            button.setMinimumHeight(26)
            button.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
            job_buttons.addWidget(button, 2 if button is self.start_button else 1)
        self._action_layout.addWidget(button_container, 1)
        root_layout.addWidget(self.action_bar)
        self.setCentralWidget(central)
        self._update_responsive_layout()

    @staticmethod
    def _section(step: str, title: str, description: str) -> tuple[QFrame, QVBoxLayout]:
        panel = QFrame()
        panel.setObjectName("Surface")
        panel.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred)
        layout = QVBoxLayout(panel)
        layout.setContentsMargins(20, 18, 20, 20)
        layout.setSpacing(8)
        title_row = QHBoxLayout()
        title_row.setSpacing(9)
        if step:
            step_label = QLabel(step)
            step_label.setObjectName("Step")
            step_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
            step_label.setFixedSize(24, 24)
            title_row.addWidget(step_label)
        title_label = QLabel(title)
        title_label.setObjectName("SectionTitle")
        title_row.addWidget(title_label, 1)
        layout.addLayout(title_row)
        if description:
            helper = PlainTextLabel(description)
            helper.setObjectName("Helper")
            layout.addWidget(helper)
        layout.addSpacing(6)
        return panel, layout

    @staticmethod
    def _field_label(text: str, field: QWidget) -> QLabel:
        label = QLabel(text)
        label.setObjectName("FieldLabel")
        label.setBuddy(field)
        field.setAccessibleName(text)
        return label

    @staticmethod
    def _text_view(label: QLabel, height: int) -> QScrollArea:
        area = QScrollArea()
        area.setObjectName("ValueScroll")
        area.setWidgetResizable(True)
        area.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        area.setFixedHeight(height)
        content = QWidget()
        content.setObjectName("ValueContent")
        layout = QVBoxLayout(content)
        layout.setContentsMargins(12, 10, 12, 10)
        layout.addWidget(label)
        label.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignTop)
        area.setWidget(content)
        return area

    def resizeEvent(self, event: QResizeEvent) -> None:
        super().resizeEvent(event)
        if hasattr(self, "_content_layout"):
            self._update_responsive_layout()

    def _update_responsive_layout(self) -> None:
        wide = self.width() >= 1040
        self._content_layout.setDirection(
            QBoxLayout.Direction.LeftToRight if wide else QBoxLayout.Direction.TopToBottom
        )
        self._content_layout.setStretch(0, 3 if wide else 0)
        self._content_layout.setStretch(1, 2 if wide else 0)
        self._action_layout.setDirection(
            QBoxLayout.Direction.LeftToRight if wide else QBoxLayout.Direction.TopToBottom
        )

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
            EXCEL_FILE_FILTER,
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
        tone = {
            AppState.READY.value: "success",
            AppState.COMPLETED.value: "success",
            AppState.INPUTTING.value: "active",
            AppState.WAITING_FOR_CLEAR.value: "active",
            AppState.LOADING_EXCEL.value: "active",
            AppState.PAUSED.value: "warning",
            AppState.STOPPING.value: "warning",
            AppState.ERROR.value: "error",
        }.get(state, "idle")
        set_style_property(self.state_value_label, "tone", tone)
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
        self.pause_button.setVisible(state != AppState.PAUSED.value)
        self.resume_button.setVisible(state == AppState.PAUSED.value)
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
        self.selector_label.setVisible(is_text_mode)
        self.selector_edit.setVisible(is_text_mode)
        self.training_url_label.setVisible(not is_text_mode)
        self.training_url_edit.setVisible(not is_text_mode)
        self.audit_notice_label.setVisible(not is_text_mode)
        self.test_selector_button.setText(
            "Text 영역 확인" if is_text_mode else "AWS 페이지 확인"
        )
        self.test_selector_button.setToolTip(
            "현재 페이지에서 입력 가능한 요소가 정확히 하나인지 확인합니다."
            if is_text_mode else
            "로그인 후 교육 상세 페이지로 이동한 다음 눌러주세요. 할당 대상 과정을 고정합니다."
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
        set_style_property(
            self.ready_indicator_label, "ready",
            "true" if self._aws_page_ready and not is_text_mode else "false",
        )
        self.ready_indicator_label.setText(
            "입력 후 웹 Action은 직접 실행합니다. 빈 상태가 유지되면 다음 값을 입력합니다."
            if is_text_mode else (
                "과정 확인 완료 · 확인된 과정 URL을 확인한 후 시작하세요."
                if self._aws_page_ready else
                "과정 상세페이지에 도착하면 AWS 페이지 확인을 눌러주세요."
            )
        )
        self.guide_label.setText(
            "1. Excel 파일 · Sheet · 시작 셀을 선택하세요.\n"
            "2. 브라우저를 열고 입력 영역을 확인하세요.\n"
            "3. 시작 후 웹 Action은 직접 실행하세요."
            if is_text_mode else
            "1. Excel 파일 · Sheet · 시작 셀을 선택하세요.\n"
            "2. 브라우저에서 로그인 후 과정으로 이동하세요.\n"
            "3. AWS 페이지 확인 후 시작을 누르세요."
        )
        state = self.state_value_label.text()
        hints = {
            AppState.IDLE.value: "Excel 파일을 선택해 작업을 준비하세요.",
            AppState.LOADING_EXCEL.value: "Excel 데이터를 불러오는 중입니다.",
            AppState.INPUTTING.value: "작업 중입니다. 일시정지하거나 안전하게 중지할 수 있습니다.",
            AppState.WAITING_FOR_CLEAR.value: "현재 작업의 완료를 기다리고 있습니다.",
            AppState.PAUSED.value: "일시정지됨 · 재개하면 현재 단계부터 이어집니다.",
            AppState.STOPPING.value: "안전하게 중지하는 중입니다. 잠시 기다려 주세요.",
            AppState.COMPLETED.value: "처리가 끝났습니다. 결과를 확인한 후 다음 작업을 준비하세요.",
            AppState.ERROR.value: "마지막 메시지를 확인하세요. 재시작 전 현재 처리 결과를 확인해 주세요.",
        }
        self.action_hint_label.setText(hints.get(state, (
            "준비가 완료되었습니다. 시작 버튼을 눌러주세요."
            if is_text_mode or self._aws_page_ready else
            "수동 로그인 후 교육 상세페이지에서 AWS 페이지 확인을 눌러주세요."
        )))

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

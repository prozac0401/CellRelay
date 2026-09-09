"""UI-only regression coverage; no workbook, browser, or AWS calls occur."""

from __future__ import annotations

import pytest
from PySide6.QtCore import QObject, QPoint, QPointF, Qt, Signal
from PySide6.QtGui import QWheelEvent
from PySide6.QtTest import QTest
from PySide6.QtWidgets import QMessageBox, QPushButton, QScrollArea

from app.config.settings import AWS_LOGIN_URL, AppSettings
from app.core.state import AppState, ProgressSnapshot
from app.ui.main_window import MainWindow


class FakeController(QObject):
    """Small signal-compatible port, so UI tests cannot submit real users."""

    state_changed = Signal(str)
    sheets_changed = Signal(list)
    progress_changed = Signal(object)
    message_changed = Signal(str)
    target_ready_changed = Signal(bool)
    training_confirmed = Signal(str)
    next_start_cell_changed = Signal(str)
    shutdown_completed = Signal()

    def __init__(self):
        super().__init__()
        self.settings = AppSettings()
        self.state = AppState.IDLE
        self.progress = ProgressSnapshot()
        self.aws_page_ready = False
        self.calls = []

    def load_excel(self, path):
        self.calls.append(("load_excel", path))

    def select_sheet(self, name):
        self.calls.append(("select_sheet", name))

    def open_browser(self, url):
        self.calls.append(("open_browser", url))

    def test_target(self, workflow_mode, selector, url):
        self.calls.append(("test_target", workflow_mode, selector, url))

    def start_job(self, **settings):
        self.calls.append(("start_job", settings))

    def pause_job(self):
        self.calls.append(("pause_job",))

    def resume_job(self):
        self.calls.append(("resume_job",))

    def stop_job(self):
        self.calls.append(("stop_job",))

    def shutdown(self):
        self.calls.append(("shutdown",))

    def set_state(self, state):
        self.state = state
        self.state_changed.emit(state.value)

    def set_target_ready(self, ready):
        self.aws_page_ready = ready
        self.target_ready_changed.emit(ready)

    def set_progress(self, **values):
        self.progress = ProgressSnapshot(**values)
        self.progress_changed.emit(self.progress)


def settle(app):
    """Allow resize, layout and deferred UI signals to be delivered."""
    for _ in range(5):
        app.processEvents()


@pytest.fixture
def window(qapp):
    controller = FakeController()
    view = MainWindow(controller)
    view.show()
    settle(qapp)
    yield view, controller
    view._can_close = True
    view.close()
    view.deleteLater()
    settle(qapp)


def test_idle_is_safe_and_uses_login_url(window):
    view, controller = window
    assert view.state_value_label.text() == "IDLE"
    assert view.url_edit.text() == AWS_LOGIN_URL
    assert view.start_cell_edit.text() == "B5"
    assert view.workflow_combo.currentData() == "aws_skill_builder"
    assert view.browse_button.isEnabled()
    assert view.open_browser_button.isEnabled()
    assert view.test_selector_button.isEnabled()
    for button in (
        view.start_button,
        view.pause_button,
        view.resume_button,
        view.stop_button,
    ):
        assert not button.isEnabled()
    assert not controller.calls


@pytest.mark.parametrize("state", [AppState.READY, AppState.ERROR, AppState.COMPLETED])
def test_aws_start_requires_confirmed_target(window, state):
    view, controller = window
    controller.set_state(state)
    assert not view.start_button.isEnabled()
    controller.set_target_ready(True)
    assert view.start_button.isEnabled()
    controller.set_target_ready(False)
    assert not view.start_button.isEnabled()


def test_modes_show_only_relevant_target_settings(window, qapp):
    view, controller = window
    controller.set_state(AppState.READY)
    assert not view.selector_edit.isVisible()
    assert view.training_url_edit.isVisible()
    assert view.test_selector_button.text() == "AWS 페이지 확인"

    view.workflow_combo.setCurrentIndex(view.workflow_combo.findData("text_clear"))
    settle(qapp)
    assert view.selector_edit.isVisible()
    assert view.selector_edit.isEnabled()
    assert not view.training_url_edit.isVisible()
    assert view.test_selector_button.text() == "Text 영역 확인"
    assert view.start_button.isEnabled()

    view.workflow_combo.setCurrentIndex(
        view.workflow_combo.findData("aws_skill_builder")
    )
    assert not view.selector_edit.isVisible()
    assert not view.start_button.isEnabled()


@pytest.mark.parametrize(
    ("state", "can_pause", "can_resume", "can_stop"),
    [
        (AppState.INPUTTING, True, False, True),
        (AppState.WAITING_FOR_CLEAR, True, False, True),
        (AppState.PAUSED, False, True, True),
        (AppState.STOPPING, False, False, False),
    ],
)
def test_active_state_controls_remain_safe(
    window, state, can_pause, can_resume, can_stop
):
    view, controller = window
    controller.set_target_ready(True)
    controller.set_state(state)
    assert view.state_value_label.text() == state.value
    for control in (
        view.browse_button,
        view.sheet_combo,
        view.start_cell_edit,
        view.url_edit,
        view.selector_edit,
        view.workflow_combo,
        view.open_browser_button,
        view.test_selector_button,
        view.start_button,
    ):
        assert not control.isEnabled()
    assert view.pause_button.isEnabled() is can_pause
    assert view.resume_button.isEnabled() is can_resume
    assert view.stop_button.isEnabled() is can_stop


def test_loading_excel_cannot_start_even_with_previous_confirmation(window):
    view, controller = window
    controller.set_target_ready(True)
    controller.set_state(AppState.LOADING_EXCEL)
    assert not view.start_button.isEnabled()
    assert not view.pause_button.isEnabled()
    assert not view.resume_button.isEnabled()
    assert not view.stop_button.isEnabled()


def test_progress_and_messages_are_literal_selectable_text(window, qapp):
    view, controller = window
    value = "<b>literal Excel text</b> & " + "verylongidentifier" * 40
    message = "<a href='https://example.invalid'>not a link</a> " + "메시지 " * 90
    controller.set_progress(
        current_cell="B17",
        current_value=value,
        processed_count=13,
        skipped_count=2,
        total_items=120,
        last_message=message,
    )
    settle(qapp)
    assert view.cell_value_label.text() == "B17"
    assert view.current_value_label.text() == value
    assert view.message_value_label.text() == message
    for label in (view.current_value_label, view.message_value_label):
        assert label.textFormat() == Qt.TextFormat.PlainText
        assert (
            label.textInteractionFlags() & Qt.TextInteractionFlag.TextSelectableByMouse
        )
    assert view.count_value_label.text() == "13 / 120"
    assert view.skipped_value_label.text() == "2"
    assert view.progress_bar.value() == 13
    assert view.progress_bar.maximum() == 120
    controller.message_changed.emit("<b>still literal</b>")
    assert view.message_value_label.text() == "<b>still literal</b>"
    assert view.message_value_label.textFormat() == Qt.TextFormat.PlainText


def test_unknown_progress_does_not_show_spurious_completion(window):
    view, controller = window
    controller.set_progress(processed_count=0, total_items=0)
    assert view.progress_bar.value() == 0
    assert view.progress_bar.maximum() == 1
    assert "?" in view.progress_bar.format()


@pytest.mark.parametrize(
    "size",
    [
        (560, 320),
        (560, 340),
        (780, 420),
        (560, 560),
        (780, 660),
        (1100, 800),
        (1500, 1000),
    ],
)
@pytest.mark.parametrize("mode", ["aws_skill_builder", "text_clear"])
def test_responsive_window_keeps_controls_reachable_without_horizontal_scroll(
    window, qapp, size, mode
):
    view, controller = window
    view.workflow_combo.setCurrentIndex(view.workflow_combo.findData(mode))
    controller.set_progress(
        current_cell="B17000",
        current_value="long-identifier-" * 100,
        last_message="긴 오류 메시지와 상세 설명 " * 100,
    )
    view.url_edit.setText("https://skillbuilder.aws/" + "very-long-path/" * 80)
    view.resize(*size)
    settle(qapp)
    assert view.width() == size[0]
    assert view.height() == size[1]

    for scroll in view.findChildren(QScrollArea):
        assert scroll.horizontalScrollBar().maximum() == 0

    essential = [
        view.excel_path_edit,
        view.browse_button,
        view.sheet_combo,
        view.start_cell_edit,
        view.url_edit,
        view.workflow_combo,
        view.open_browser_button,
        view.test_selector_button,
        view.start_button,
        view.stop_button,
        view.progress_bar,
    ]
    essential.append(
        view.selector_edit if mode == "text_clear" else view.training_url_edit
    )
    for control in essential:
        assert control.isVisible(), control.objectName()
        point = control.mapTo(view, QPoint(0, 0))
        assert point.x() >= 0
        assert point.x() + control.width() <= view.width()
        assert control.width() > 20
        assert control.height() > 0
        if isinstance(control, QPushButton):
            assert control.width() >= control.fontMetrics().horizontalAdvance(
                control.text()
            ) + 12
        # Offscreen vertically is acceptable only if a real scroll viewport
        # can reveal the control. The action bar remains pinned and visible.
        ancestor = control.parentWidget()
        while ancestor is not None and not isinstance(ancestor, QScrollArea):
            ancestor = ancestor.parentWidget()
        if isinstance(ancestor, QScrollArea):
            ancestor.ensureWidgetVisible(control)
            settle(qapp)
            center = control.mapTo(ancestor.viewport(), control.rect().center())
            assert ancestor.viewport().rect().contains(center)
        else:
            assert point.y() >= 0
            assert point.y() + control.height() <= view.height()


def test_layout_reflows_in_both_directions_and_keeps_actions_pinned(window, qapp):
    view, controller = window
    controller.set_state(AppState.WAITING_FOR_CLEAR)
    for width in (1500, 560, 1100, 780):
        view.resize(width, 800)
        settle(qapp)
        settings_pos = view.settings_panel.mapTo(view, QPoint(0, 0))
        status_pos = view.status_panel.mapTo(view, QPoint(0, 0))
        if width >= 1040:
            assert settings_pos.x() < status_pos.x()
            assert abs(settings_pos.y() - status_pos.y()) <= 1
        else:
            assert settings_pos.y() < status_pos.y()
            assert abs(settings_pos.x() - status_pos.x()) <= 1
        before = view.stop_button.mapTo(view, QPoint(0, 0))
        scrollbar = view.content_scroll.verticalScrollBar()
        scrollbar.setValue(scrollbar.maximum())
        settle(qapp)
        assert view.stop_button.mapTo(view, QPoint(0, 0)) == before
        assert view.stop_button.isEnabled()
        assert before.y() + view.stop_button.height() <= view.height()


def test_initial_window_fits_available_screen_where_minimum_fits(window, qapp):
    view, _ = window
    settle(qapp)
    available = view.screen().availableGeometry()
    if view.minimumWidth() <= available.width():
        assert view.width() <= available.width()
    if view.minimumHeight() <= available.height():
        assert view.height() <= available.height()


@pytest.mark.parametrize("width", [560, 1100])
def test_large_excel_row_and_counts_do_not_force_horizontal_overflow(
    window, qapp, width
):
    view, controller = window
    controller.set_progress(
        current_cell="XFD1048576",
        processed_count=1048576,
        skipped_count=1048576,
        total_items=1048576,
    )
    view.resize(width, 800)
    settle(qapp)
    assert view.width() == width
    assert view.content_scroll.horizontalScrollBar().maximum() == 0
    assert view.cell_value_label.text() == "XFD1048576"
    assert view.count_value_label.text() == "1048576 / 1048576"
    assert view.skipped_value_label.text() == "1048576"


def test_sheet_and_target_signals_update_without_extra_controller_actions(window):
    view, controller = window
    controller.settings.sheet = "Second"
    controller.sheets_changed.emit(["First", "Second"])
    assert view.sheet_combo.currentText() == "Second"
    assert not controller.calls
    controller.next_start_cell_changed.emit("B18")
    assert view.start_cell_edit.text() == "B18"
    training = "https://skillbuilder.aws/admin/training/example"
    controller.training_confirmed.emit(training)
    assert view.training_url_edit.text() == training


def test_browser_and_target_buttons_forward_only_user_intent(window):
    view, controller = window
    view.open_browser_button.click()
    view.test_selector_button.click()
    assert controller.calls == [
        ("open_browser", AWS_LOGIN_URL),
        ("test_target", "aws_skill_builder", "textarea", AWS_LOGIN_URL),
    ]


@pytest.mark.parametrize("combo_name", ["sheet_combo", "workflow_combo"])
@pytest.mark.parametrize("wheel_delta", [-120, 120])
@pytest.mark.parametrize("focused", [False, True])
def test_scrolling_settings_does_not_change_closed_combo_selection(
    window, qapp, combo_name, wheel_delta, focused
):
    view, controller = window
    controller.sheets_changed.emit(["First", "Second", "Third"])
    combo = getattr(view, combo_name)
    combo.setCurrentIndex(1 if wheel_delta > 0 else 0)
    controller.calls.clear()
    (combo if focused else view.start_cell_edit).setFocus()
    settle(qapp)
    original_index = combo.currentIndex()
    original_mode = view.workflow_combo.currentData()
    assert not combo.view().isVisible()
    event = QWheelEvent(
        QPointF(combo.rect().center()),
        QPointF(combo.mapToGlobal(combo.rect().center())),
        QPoint(0, 0),
        QPoint(0, wheel_delta),
        Qt.MouseButton.NoButton,
        Qt.KeyboardModifier.NoModifier,
        Qt.ScrollPhase.NoScrollPhase,
        False,
    )
    qapp.sendEvent(combo, event)
    settle(qapp)
    assert combo.currentIndex() == original_index
    assert view.workflow_combo.currentData() == original_mode
    assert not controller.calls
    assert not event.isAccepted()


@pytest.mark.parametrize("combo_name", ["sheet_combo", "workflow_combo"])
def test_combo_keyboard_selection_remains_available(window, qapp, combo_name):
    view, controller = window
    controller.sheets_changed.emit(["First", "Second"])
    combo = getattr(view, combo_name)
    combo.setCurrentIndex(0)
    combo.setFocus()
    settle(qapp)
    QTest.keyClick(combo, Qt.Key.Key_Down)
    settle(qapp)
    assert combo.currentIndex() == 1
    QTest.keyClick(combo, Qt.Key.Key_Up)
    settle(qapp)
    assert combo.currentIndex() == 0


def test_aws_start_requires_user_acceptance_before_forwarding(window, monkeypatch):
    view, controller = window
    controller.set_state(AppState.READY)
    controller.set_target_ready(True)
    controller.sheets_changed.emit(["Data"])
    monkeypatch.setattr(
        QMessageBox, "question", lambda *args: QMessageBox.StandardButton.No
    )
    view.start_button.click()
    assert not controller.calls
    monkeypatch.setattr(
        QMessageBox, "question", lambda *args: QMessageBox.StandardButton.Yes
    )
    view.start_button.click()
    assert controller.calls == [
        (
            "start_job",
            {
                "file_path": "",
                "sheet": "Data",
                "start_cell": "B5",
                "url": AWS_LOGIN_URL,
                "selector": "textarea",
                "workflow_mode": "aws_skill_builder",
            },
        )
    ]


def test_close_waits_for_shutdown_ack_and_does_not_repeat_shutdown(window, qapp):
    view, controller = window
    controller.set_state(AppState.INPUTTING)
    view.close()
    settle(qapp)
    assert view.isVisible()
    assert not view.centralWidget().isEnabled()
    assert controller.calls == [("shutdown",)]
    view.close()
    assert controller.calls == [("shutdown",)]
    controller.shutdown_completed.emit()
    settle(qapp)
    assert not view.isVisible()

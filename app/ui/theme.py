"""Shared visual tokens for the native Qt UI; no workflow behavior lives here."""

from __future__ import annotations

from PySide6.QtGui import QFont
from PySide6.QtWidgets import QWidget


STYLESHEET = """
QMainWindow, QWidget#AppRoot { background: #F4F6FA; color: #19243B; }
QWidget#ScrollContent { background: #F4F6FA; }
QWidget { font-family: "Malgun Gothic", "Segoe UI"; font-size: 13px; }
QLabel { color: #19243B; background: transparent; border: none; }
QLabel#Brand { font-size: 26px; font-weight: 700; letter-spacing: -1px; }
QLabel#BrandMark { background: #245FE5; color: white; border-radius: 12px;
    font-size: 19px; font-weight: 700; }
QLabel#Version { color: #728098; font-size: 12px; }
QLabel#Subtitle, QLabel#Helper, QLabel#FieldLabel, QLabel#MetricLabel {
    color: #65738A;
}
QLabel#Subtitle { font-size: 12px; }
QLabel#SectionTitle { font-size: 16px; font-weight: 700; }
QLabel#Step { color: #245FE5; background: #ECF2FF; border-radius: 10px;
    font-size: 12px; font-weight: 700; }
QLabel#Helper, QLabel#MetricLabel { font-size: 12px; }
QLabel#MetricValue { font-size: 23px; font-weight: 700; }
QLabel#CellValue { font-size: 29px; font-weight: 700; color: #245FE5; }
QLabel#StateBadge { border-radius: 9px; padding: 5px 9px;
    background: #EDF1F6; color: #52617A; font-size: 10px; font-weight: 700; }
QLabel#StateBadge[tone="active"] { background: #EAF1FF; color: #2057C8; }
QLabel#StateBadge[tone="success"] { background: #E8F6EF; color: #24714D; }
QLabel#StateBadge[tone="warning"] { background: #FFF4DF; color: #896015; }
QLabel#StateBadge[tone="error"] { background: #FFF0F0; color: #B84040; }
QLabel#ReadyIndicator { font-size: 12px; color: #65738A; }
QLabel#ReadyIndicator[ready="true"] { color: #24714D; }
QLabel#ActionHint { color: #65738A; font-size: 12px; }
QFrame#Surface { background: white; border: 1px solid #E1E6EF; border-radius: 14px; }
QFrame#ActionBar { background: white; border: none; border-top: 1px solid #E1E6EF; }
QFrame#Divider { background: #E9EDF3; border: none; }
QLineEdit, QComboBox { background: #FFFFFF; color: #19243B;
    border: 1px solid #D6DDE8; border-radius: 7px; padding: 8px 10px;
    min-height: 18px; selection-background-color: #D9E7FF; selection-color: #19243B; }
QLineEdit:hover, QComboBox:hover { border-color: #A6B5CC; }
QLineEdit:focus, QComboBox:focus { border: 2px solid #5082EE; padding: 7px 9px; }
QLineEdit:read-only { background: #F7F9FC; color: #53627B; }
QLineEdit:disabled, QComboBox:disabled { color: #7D8799; background: #F3F5F8;
    border-color: #E4E8EF; }
QComboBox { padding-right: 28px; }
QComboBox:focus { padding-right: 27px; }
QComboBox::drop-down { border: none; width: 26px; }
QComboBox::down-arrow { width: 0; height: 0; }
QComboBox QAbstractItemView { background: white; color: #19243B;
    selection-background-color: #EAF1FF; selection-color: #245FE5;
    border: 1px solid #D6DDE8; padding: 4px; outline: none; }
QPushButton { background: #FFFFFF; color: #32435E; border: 1px solid #D6DDE8;
    border-radius: 7px; padding: 8px 15px; min-height: 18px; font-weight: 600; }
QPushButton:hover { background: #F1F5FC; border-color: #A8BAD8; }
QPushButton:pressed { background: #E8EFFB; }
QPushButton:focus { border: 2px solid #5082EE; padding: 7px 14px; }
QPushButton#PrimaryButton { background: #245FE5; color: white; border-color: #245FE5; }
QPushButton#PrimaryButton:hover { background: #1D52CC; border-color: #1D52CC; }
QPushButton#PrimaryButton:pressed { background: #1747B3; }
QPushButton#PrimaryButton:focus { border-color: #163F9B; }
QPushButton#DangerButton { color: #AD3C44; border-color: #E5C9CE; }
QPushButton#DangerButton:hover { background: #FFF2F3; border-color: #D7A4AB; }
QPushButton:disabled, QPushButton#PrimaryButton:disabled, QPushButton#DangerButton:disabled {
    color: #929BAD; background: #F1F3F7; border-color: #E6EAF0; }
QScrollArea { border: none; background: transparent; }
QScrollArea#ValueScroll { background: #F7F9FC; border: 1px solid #E8EDF4; border-radius: 8px; }
QWidget#ValueContent { background: #F7F9FC; }
QScrollBar:vertical { background: transparent; width: 9px; margin: 2px 1px; }
QScrollBar::handle:vertical { background: #C4CDDA; border-radius: 3px; min-height: 28px; }
QScrollBar::handle:vertical:hover { background: #9CAAC0; }
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical { height: 0; }
QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical { background: transparent; }
QProgressBar { background: #EAF0F8; border: none; border-radius: 4px; height: 8px; }
QProgressBar::chunk { background: #3974EF; border-radius: 4px; }
QToolTip { background: #24334C; color: white; border: none; padding: 6px; }
"""


def apply_theme(widget: QWidget) -> None:
    """Scope styling to this window; respect Qt's device-independent layout."""
    widget.setFont(QFont("Malgun Gothic", 10))
    widget.setStyleSheet(STYLESHEET)


def set_style_property(widget: QWidget, name: str, value: str) -> None:
    """Refresh selectors when a state-specific visual property changes."""
    if widget.property(name) == value:
        return
    widget.setProperty(name, value)
    widget.style().unpolish(widget)
    widget.style().polish(widget)
    widget.update()

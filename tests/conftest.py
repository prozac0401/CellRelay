"""Keep one GUI-capable Qt application for both controller and window tests.

The controller tests need only QCoreApplication, but Qt cannot upgrade that
instance to QApplication later. Creating it first also keeps UI tests headless
and independent of an open CellRelay/Edge window on the developer's desktop.
"""

import os
from pathlib import Path

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

import pytest
from PySide6.QtGui import QFontDatabase
from PySide6.QtWidgets import QApplication


@pytest.fixture(scope="session", autouse=True)
def qapp():
    app = QApplication.instance() or QApplication([])
    assert isinstance(app, QApplication)
    app.setQuitOnLastWindowClosed(False)
    # Windows' offscreen backend does not populate the system font database.
    # Register both weights so wrapping and button metrics match the actual UI
    # rather than passing layout assertions against missing-glyph squares.
    font_dir = Path(os.environ.get("WINDIR", "C:/Windows")) / "Fonts"
    for font_name in ("malgun.ttf", "malgunbd.ttf"):
        font_path = font_dir / font_name
        if font_path.is_file():
            assert QFontDatabase.addApplicationFont(str(font_path)) >= 0
    yield app
    app.processEvents()

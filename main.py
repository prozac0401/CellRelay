"""CellRelay desktop application entry point."""

from __future__ import annotations

import logging
import os
import sys
from pathlib import Path

from PySide6.QtCore import QTimer
from PySide6.QtWidgets import QApplication

from app.core.controller import CellRelayController
from app.core.logging_config import configure_logging
from app.ui.main_window import MainWindow


def _run_playwright_smoke_test() -> int:
    """Verify that a packaged build can launch Edge and control a text input."""
    from playwright.sync_api import sync_playwright

    logger = logging.getLogger(__name__)
    try:
        with sync_playwright() as playwright:
            browser = playwright.chromium.launch(channel="msedge", headless=True)
            try:
                page = browser.new_page()
                page.set_content("<textarea id='target'></textarea>")
                target = page.locator("#target")
                target.fill("CellRelay smoke test")
                if target.input_value() != "CellRelay smoke test":
                    raise RuntimeError("Playwright input verification failed")
            finally:
                browser.close()
        logger.info("Packaged Playwright smoke test passed")
        return 0
    except Exception:
        logger.exception("Packaged Playwright smoke test failed")
        return 1


def _application_data_root() -> Path:
    """Use a persistent writable directory when running as a packaged EXE."""
    if getattr(sys, "frozen", False):
        local_app_data = os.environ.get("LOCALAPPDATA")
        base = (
            Path(local_app_data)
            if local_app_data
            else Path.home() / "AppData" / "Local"
        )
        root = base / "CellRelay"
        root.mkdir(parents=True, exist_ok=True)
        return root
    return Path(__file__).resolve().parent


def main() -> int:
    smoke_test = "--smoke-test" in sys.argv
    playwright_smoke_test = "--playwright-smoke-test" in sys.argv
    internal_options = {"--smoke-test", "--playwright-smoke-test"}
    application_args = [arg for arg in sys.argv if arg not in internal_options]
    data_root = _application_data_root()
    configure_logging(data_root / "logs")
    logger = logging.getLogger(__name__)
    logger.info("Application started")
    if playwright_smoke_test:
        exit_code = _run_playwright_smoke_test()
        logger.info("Application stopped")
        return exit_code

    application = QApplication(application_args)
    application.setApplicationName("CellRelay")
    application.setOrganizationName("CellRelay")
    controller = CellRelayController(data_root)
    window = MainWindow(controller)
    window.show()
    if smoke_test:
        QTimer.singleShot(750, window.close)

    exit_code = application.exec()
    controller.shutdown()
    if not controller.wait_for_shutdown():
        logger.warning("Browser worker did not stop within the shutdown timeout")
    logger.info("Application stopped")
    return exit_code


if __name__ == "__main__":
    raise SystemExit(main())

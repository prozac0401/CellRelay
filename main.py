"""CellRelay desktop application entry point."""

from __future__ import annotations

import logging
import os
import sys
import tempfile
from pathlib import Path

from PySide6.QtCore import QTimer
from PySide6.QtWidgets import QApplication

from app import __version__
from app.core.controller import CellRelayController
from app.core.logging_config import configure_logging
from app.ui.main_window import MainWindow


def _run_playwright_smoke_test() -> int:
    """Verify that a packaged build can launch Edge and control a text input."""
    import asyncio

    from playwright.async_api import async_playwright

    logger = logging.getLogger(__name__)

    async def run() -> None:
        async with async_playwright() as playwright:
            browser = await playwright.chromium.launch(
                channel="msedge",
                headless=True,
            )
            try:
                page = await browser.new_page()
                await page.set_content("<textarea id='target'></textarea>")
                target = page.locator("#target")
                await target.fill("CellRelay smoke test")
                if await target.input_value() != "CellRelay smoke test":
                    raise RuntimeError("Playwright input verification failed")
            finally:
                await browser.close()

    try:
        asyncio.run(run())
        logger.info("Packaged Playwright smoke test passed")
        return 0
    except Exception:
        logger.exception("Packaged Playwright smoke test failed")
        return 1


def _run_browser_worker_smoke_test() -> int:
    """Exercise two queued Playwright calls on the production QThread path."""
    logger = logging.getLogger(__name__)
    application = QApplication([sys.argv[0], "--browser-worker-smoke-test"])
    application.setApplicationName("CellRelay")
    application.setOrganizationName("CellRelay")
    smoke_root = Path(tempfile.gettempdir()) / f"CellRelay-worker-smoke-{os.getpid()}"
    controller = CellRelayController(smoke_root)
    result = {"done": False, "success": False}

    def finish(success: bool, message: str) -> None:
        if result["done"]:
            return
        result["done"] = True
        result["success"] = success
        if success:
            logger.info("Browser worker QThread smoke test passed")
        else:
            logger.error("Browser worker QThread smoke test failed: %s", message)
        application.quit()

    def on_browser_status(success: bool, message: str) -> None:
        if not success:
            finish(False, message)
            return
        controller.test_target("text_clear", "#target")

    controller.browser_status_changed.connect(on_browser_status)
    controller.selector_test_result.connect(finish)
    QTimer.singleShot(30_000, lambda: finish(False, "30초 제한 시간 초과"))
    QTimer.singleShot(
        0,
        lambda: controller.open_browser(
            "data:text/html,<textarea%20id='target'></textarea>"
        ),
    )
    application.exec()
    controller.shutdown()
    if not controller.wait_for_shutdown(10_000):
        logger.error("Browser worker did not stop after QThread smoke test")
        return 1
    return 0 if result["success"] else 1


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
    browser_worker_smoke_test = "--browser-worker-smoke-test" in sys.argv
    excel_smoke_test = "--excel-smoke-test" in sys.argv
    internal_options = {
        "--smoke-test",
        "--playwright-smoke-test",
        "--browser-worker-smoke-test",
        "--excel-smoke-test",
    }
    application_args = [arg for arg in sys.argv if arg not in internal_options]
    data_root = _application_data_root()
    configure_logging(data_root / "logs")
    logger = logging.getLogger(__name__)
    logger.info(
        "Application started version=%s executable=%s data_root=%s",
        __version__,
        sys.executable,
        data_root,
    )
    if excel_smoke_test:
        from app.excel.smoke import run_excel_smoke_test

        exit_code = run_excel_smoke_test()
        logger.info("Application stopped")
        return exit_code
    if playwright_smoke_test:
        exit_code = _run_playwright_smoke_test()
        logger.info("Application stopped")
        return exit_code
    if browser_worker_smoke_test:
        exit_code = _run_browser_worker_smoke_test()
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

"""Small browser integration test for the stable-empty JavaScript watcher."""

import pytest
from playwright.sync_api import Error as PlaywrightError
from playwright.sync_api import sync_playwright

from app.browser.browser_worker import (
    _WATCHER_INSTALL_SCRIPT,
    _WATCHER_STATUS_SCRIPT,
)


def test_transient_empty_does_not_complete_but_stable_empty_does() -> None:
    with sync_playwright() as playwright:
        try:
            browser = playwright.chromium.launch(channel="msedge", headless=True)
        except PlaywrightError as exc:
            pytest.skip(f"Microsoft Edge is unavailable: {exc}")

        try:
            page = browser.new_page()
            page.set_content("<textarea id='target'></textarea>")
            locator = page.locator("#target")
            locator.fill("Excel value")
            locator.evaluate(_WATCHER_INSTALL_SCRIPT, 500)

            locator.evaluate(
                """element => {
                    element.value = "";
                    element.dispatchEvent(new Event("input"));
                    setTimeout(() => {
                        element.value = "Excel value";
                        element.dispatchEvent(new Event("input"));
                    }, 200);
                }"""
            )
            page.wait_for_timeout(650)
            assert locator.evaluate(_WATCHER_STATUS_SCRIPT) == "waiting"

            locator.evaluate(
                """element => {
                    element.value = "";
                    element.dispatchEvent(new Event("input"));
                }"""
            )
            page.wait_for_function(
                '''() => document.querySelector("#target")
                    .__cellRelayClearWatcher.status === "cleared"''',
                timeout=1_500,
            )
            assert locator.evaluate(_WATCHER_STATUS_SCRIPT) == "cleared"
        finally:
            browser.close()

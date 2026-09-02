"""Playwright worker that runs entirely on a dedicated QThread."""

from __future__ import annotations

import logging
import threading
import traceback
from enum import Enum
from typing import Any

from playwright.sync_api import (
    Browser,
    BrowserContext,
    Page,
    Playwright,
    sync_playwright,
)
from PySide6.QtCore import QObject, Signal, Slot

from app.browser.aws_skill_builder import (
    AwsAssignmentCancelled,
    AwsSkillBuilderAssignment,
)

logger = logging.getLogger(__name__)


class InputMethod(str, Enum):
    """Input-strategy seam for future clipboard and keyboard support."""

    FILL = "fill"


_WATCHER_INSTALL_SCRIPT = r"""
(element, stableMs) => {
    const key = "__cellRelayClearWatcher";
    const previous = element[key];
    if (previous && typeof previous.dispose === "function") {
        previous.dispose();
    }

    const readValue = () => {
        if ("value" in element) {
            return String(element.value ?? "");
        }
        return String(element.textContent ?? "");
    };

    const watcher = {
        status: "waiting",
        emptySince: null,
        timer: null,
        interval: null,
        observer: null,
        disposed: false,
    };

    const cancelTimer = () => {
        if (watcher.timer !== null) {
            clearTimeout(watcher.timer);
            watcher.timer = null;
        }
    };

    const cleanup = () => {
        cancelTimer();
        if (watcher.interval !== null) {
            clearInterval(watcher.interval);
            watcher.interval = null;
        }
        if (watcher.observer !== null) {
            watcher.observer.disconnect();
        }
        element.removeEventListener("input", check);
        element.removeEventListener("change", check);
    };

    const check = () => {
        if (watcher.disposed || watcher.status !== "waiting") {
            return;
        }
        if (readValue() !== "") {
            watcher.emptySince = null;
            cancelTimer();
            return;
        }
        if (watcher.emptySince !== null) {
            return;
        }
        watcher.emptySince = performance.now();
        watcher.timer = setTimeout(() => {
            watcher.timer = null;
            if (!watcher.disposed && readValue() === "") {
                watcher.status = "cleared";
                watcher.clearedAt = Date.now();
                cleanup();
            } else {
                watcher.emptySince = null;
                check();
            }
        }, stableMs);
    };

    watcher.dispose = () => {
        watcher.disposed = true;
        watcher.status = "disposed";
        cleanup();
    };

    element[key] = watcher;
    element.addEventListener("input", check);
    element.addEventListener("change", check);
    watcher.observer = new MutationObserver(check);
    watcher.observer.observe(element, {
        attributes: true,
        childList: true,
        subtree: true,
        characterData: true,
    });

    // Events and MutationObserver are primary. This low-frequency fallback also
    // catches frameworks that assign the DOM value property without an event.
    const sampleMs = Math.min(200, Math.max(50, Math.floor(stableMs / 4)));
    watcher.interval = setInterval(check, sampleMs);
    check();
    return watcher.status;
}
"""

_WATCHER_STATUS_SCRIPT = r"""
element => element.__cellRelayClearWatcher?.status ?? "missing"
"""

_WATCHER_DISPOSE_SCRIPT = r"""
element => {
    const watcher = element.__cellRelayClearWatcher;
    if (watcher && typeof watcher.dispose === "function") {
        watcher.dispose();
    }
}
"""

_ELEMENT_INFO_SCRIPT = r"""
element => ({
    tag: element.tagName.toLowerCase(),
    type: String(element.getAttribute("type") ?? "text").toLowerCase(),
    disabled: Boolean(element.disabled),
    readOnly: Boolean(element.readOnly),
    contentEditable: element.getAttribute("contenteditable"),
})
"""


class BrowserWorker(QObject):
    """Own Playwright resources and serialize browser operations in one thread."""

    browser_opened = Signal(str, str)
    selector_tested = Signal(bool, str, int)
    aws_page_tested = Signal(bool, str, str)
    text_inputted = Signal(int, str)
    clear_detected = Signal(int)
    assignment_stage_changed = Signal(int, str, str)
    assignment_completed = Signal(int)
    operation_cancelled = Signal(int, str)
    operation_failed = Signal(str, int, str, str)
    shutdown_finished = Signal()

    def __init__(self) -> None:
        super().__init__()
        self._playwright: Playwright | None = None
        self._browser: Browser | None = None
        self._context: BrowserContext | None = None
        self._page: Page | None = None
        self._browser_name = ""
        self._stop_event = threading.Event()
        self._pause_event = threading.Event()

    def reset_control_flags(self) -> None:
        """Thread-safe control call made before a new run is queued."""
        self._stop_event.clear()
        self._pause_event.clear()

    def request_pause(self) -> None:
        """Thread-safe pause request; no QThread event-loop turn is required."""
        self._pause_event.set()

    def request_resume(self) -> None:
        self._pause_event.clear()

    def request_stop(self) -> None:
        self._stop_event.set()
        self._pause_event.clear()

    @Slot(str, str)
    def open_browser(self, url: str, channel: str) -> None:
        try:
            self._ensure_playwright()
            if self._browser is None or not self._browser.is_connected():
                assert self._playwright is not None
                try:
                    self._browser = self._playwright.chromium.launch(
                        channel=channel or "msedge",
                        headless=False,
                    )
                    actual_browser = channel or "msedge"
                except Exception:
                    if (channel or "msedge") != "msedge":
                        raise
                    logger.warning(
                        "Microsoft Edge launch failed; trying Playwright Chromium",
                        exc_info=True,
                    )
                    self._browser = self._playwright.chromium.launch(headless=False)
                    actual_browser = "chromium"
                self._browser_name = actual_browser
                self._context = self._browser.new_context()
                self._page = self._context.new_page()
            else:
                actual_browser = self._browser_name or channel or "msedge"
                if self._page is None or self._page.is_closed():
                    assert self._context is not None
                    self._page = self._context.new_page()

            assert self._page is not None
            self._page.goto(url, wait_until="domcontentloaded", timeout=30_000)
            logger.info("Browser started (%s): %s", actual_browser, url)
            self.browser_opened.emit(url, actual_browser)
        except Exception as exc:  # noqa: BLE001 - Qt slot must report every failure.
            self._emit_failure("open_browser", 0, exc)

    @Slot(str)
    def test_selector(self, selector: str) -> None:
        try:
            locator, count = self._one_locator(selector)
            if count == 0:
                self.selector_tested.emit(
                    False,
                    "Selector에 해당하는 요소를 찾지 못했습니다.",
                    0,
                )
                return
            if count > 1:
                self.selector_tested.emit(
                    False,
                    f"Selector에 해당하는 요소가 {count}개입니다. 하나만 선택되도록 수정하세요.",
                    count,
                )
                return
            info = self._validate_fillable(locator)
            tag = info["tag"]
            self.selector_tested.emit(
                True,
                f"Text 영역을 찾았습니다. ({tag}, 1개)",
                1,
            )
            logger.info("Selector found: %s", selector)
        except Exception as exc:
            message = f"Text 영역을 확인하지 못했습니다: {exc}"
            logger.exception("Selector test failed: %s", selector)
            self.selector_tested.emit(False, message, 0)

    @Slot(str)
    def test_aws_page(self, expected_url: str) -> None:
        """Read-only login/page readiness check; never enters an Excel value."""
        try:
            workflow = AwsSkillBuilderAssignment(
                page=self._require_page(),
                stop_event=self._stop_event,
                pause_event=self._pause_event,
            )
            actual_url = workflow.validate_ready(expected_url)
            message = (
                "AWS 교육 상세 페이지와 교육 할당 버튼을 확인했습니다. "
                "이제 작업을 시작할 수 있습니다."
            )
            logger.info("AWS training page readiness check passed: %s", actual_url)
            self.aws_page_tested.emit(True, message, actual_url)
        except Exception as exc:  # noqa: BLE001 - readiness failure is user-facing.
            message = f"AWS 페이지를 확인하지 못했습니다: {exc}"
            logger.warning("AWS training page readiness check failed: %s", exc)
            self.aws_page_tested.emit(False, message, "")

    @Slot(int, str, str, str)
    def input_text(self, run_id: int, selector: str, text: str, method: str) -> None:
        try:
            if not self._wait_until_resumed(run_id, "input_text"):
                return
            if InputMethod(method) is not InputMethod.FILL:
                raise ValueError(f"지원하지 않는 입력 방식입니다: {method}")
            locator, count = self._one_locator(selector)
            if count != 1:
                raise RuntimeError(
                    f"Text 영역이 정확히 1개여야 합니다. 현재 {count}개입니다."
                )
            self._validate_fillable(locator)
            locator.fill(text, timeout=10_000)
            actual = locator.input_value(timeout=5_000)
            if actual != text:
                raise RuntimeError(
                    "입력 후 Text 영역의 값이 Excel 값과 일치하지 않습니다."
                )
            if actual == "":
                raise RuntimeError("빈 값은 입력 완료로 처리할 수 없습니다.")
            logger.info("Text inserted for run %s", run_id)
            self.text_inputted.emit(run_id, actual)
        except Exception as exc:  # noqa: BLE001 - Qt slot must report every failure.
            if self._stop_event.is_set():
                self.operation_cancelled.emit(run_id, "input_text")
                return
            self._emit_failure("input_text", run_id, exc)

    @Slot(int, str, int, int)
    def wait_for_clear(
        self,
        run_id: int,
        selector: str,
        stable_empty_ms: int,
        status_check_interval_ms: int,
    ) -> None:
        """Wait for an event-driven, stable empty value without moving Excel state."""
        try:
            locator, count = self._one_locator(selector)
            if count != 1:
                raise RuntimeError(
                    f"Text 영역이 정확히 1개여야 합니다. 현재 {count}개입니다."
                )
            locator.evaluate(_WATCHER_INSTALL_SCRIPT, max(100, stable_empty_ms))

            while True:
                if self._stop_event.is_set():
                    self._dispose_watcher(selector)
                    self.operation_cancelled.emit(run_id, "wait_for_clear")
                    return
                if self._pause_event.is_set():
                    self._stop_event.wait(0.1)
                    continue
                page = self._require_page()
                if page.is_closed():
                    raise RuntimeError("브라우저 페이지가 닫혔습니다.")

                locator = page.locator(selector)
                count = locator.count()
                if count == 1:
                    status = locator.evaluate(_WATCHER_STATUS_SCRIPT)
                    if status == "cleared":
                        logger.info("Text cleared and stable for %sms", stable_empty_ms)
                        self.clear_detected.emit(run_id)
                        return
                    if status == "missing":
                        # React can replace the element. A new stable-empty window
                        # starts on the replacement; a transient empty value is safe.
                        locator.evaluate(
                            _WATCHER_INSTALL_SCRIPT,
                            max(100, stable_empty_ms),
                        )
                elif count > 1:
                    raise RuntimeError(
                        f"감시 중 Selector가 {count}개 요소와 일치했습니다."
                    )

                interval_seconds = min(
                    0.25,
                    max(0.05, status_check_interval_ms / 1000.0),
                )
                self._stop_event.wait(interval_seconds)
        except Exception as exc:  # noqa: BLE001 - Qt slot must report every failure.
            if self._stop_event.is_set():
                self.operation_cancelled.emit(run_id, "wait_for_clear")
                return
            self._emit_failure("wait_for_clear", run_id, exc)

    @Slot(int, str, str)
    def assign_aws_user(
        self,
        run_id: int,
        search_value: str,
        expected_url: str,
    ) -> None:
        """Run one AWS-specific assignment without exposing the value in logs."""
        try:
            if not self._wait_until_resumed(run_id, "aws_assign_user"):
                return
            workflow = AwsSkillBuilderAssignment(
                page=self._require_page(),
                stop_event=self._stop_event,
                pause_event=self._pause_event,
                on_stage=lambda stage, message: self.assignment_stage_changed.emit(
                    run_id,
                    stage,
                    message,
                ),
            )
            workflow.assign_user(search_value, expected_url)
            logger.info("AWS user assignment verified for run %s", run_id)
            self.assignment_completed.emit(run_id)
        except AwsAssignmentCancelled:
            self.operation_cancelled.emit(run_id, "aws_assign_user")
        except Exception as exc:  # noqa: BLE001 - Qt slot must report every failure.
            if self._stop_event.is_set():
                self.operation_cancelled.emit(run_id, "aws_assign_user")
                return
            self._emit_failure("aws_assign_user", run_id, exc)

    @Slot()
    def shutdown(self) -> None:
        self.request_stop()
        try:
            if self._context is not None:
                self._context.close()
            if self._browser is not None:
                self._browser.close()
            if self._playwright is not None:
                self._playwright.stop()
        except Exception:
            logger.exception("Browser shutdown failed")
        finally:
            self._page = None
            self._context = None
            self._browser = None
            self._playwright = None
            self._browser_name = ""
            self.shutdown_finished.emit()

    def _ensure_playwright(self) -> None:
        if self._playwright is None:
            self._playwright = sync_playwright().start()

    def _require_page(self) -> Page:
        if self._page is None or self._page.is_closed():
            raise RuntimeError("먼저 브라우저를 여세요.")
        return self._page

    def _one_locator(self, selector: str) -> tuple[Any, int]:
        if not selector.strip():
            raise ValueError("Text Selector를 입력하세요.")
        locator = self._require_page().locator(selector)
        return locator, locator.count()

    @staticmethod
    def _validate_fillable(locator: Any) -> dict[str, Any]:
        info: dict[str, Any] = locator.evaluate(_ELEMENT_INFO_SCRIPT)
        tag = info["tag"]
        if tag not in {"textarea", "input"}:
            if info.get("contentEditable") == "true":
                raise RuntimeError("contenteditable은 현재 MVP에서 지원하지 않습니다.")
            raise RuntimeError("현재 MVP는 textarea와 input 요소만 지원합니다.")
        if tag == "input" and info["type"] not in {
            "text",
            "search",
            "email",
            "url",
            "tel",
            "password",
            "number",
        }:
            raise RuntimeError(f"fill할 수 없는 input type입니다: {info['type']}")
        if info["disabled"] or info["readOnly"]:
            raise RuntimeError("Text 영역이 disabled 또는 readonly 상태입니다.")
        if not locator.is_visible():
            raise RuntimeError("Text 영역이 현재 화면에 보이지 않습니다.")
        if not locator.is_editable():
            raise RuntimeError("Text 영역을 편집할 수 없습니다.")
        return info

    def _wait_until_resumed(self, run_id: int, operation: str) -> bool:
        while self._pause_event.is_set():
            if self._stop_event.wait(0.1):
                self.operation_cancelled.emit(run_id, operation)
                return False
        if self._stop_event.is_set():
            self.operation_cancelled.emit(run_id, operation)
            return False
        return True

    def _dispose_watcher(self, selector: str) -> None:
        try:
            page = self._require_page()
            locator = page.locator(selector)
            if locator.count() == 1:
                locator.evaluate(_WATCHER_DISPOSE_SCRIPT)
        except Exception:
            logger.debug("Could not dispose clear watcher", exc_info=True)

    def _emit_failure(self, operation: str, run_id: int, exc: Exception) -> None:
        detail = "".join(traceback.format_exception(type(exc), exc, exc.__traceback__))
        logger.error("Browser operation failed: %s", operation, exc_info=exc)
        self.operation_failed.emit(operation, run_id, str(exc), detail)

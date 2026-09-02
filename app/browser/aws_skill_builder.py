"""AWS Skill Builder-specific user-assignment workflow.

This module deliberately contains every site-specific locator so the generic
Playwright worker and the UI remain reusable for other pages.
"""

from __future__ import annotations

import re
import threading
import time
from collections.abc import Callable
from dataclasses import dataclass
from urllib.parse import parse_qs, urlparse

from playwright.sync_api import Locator, Page

_TRAINING_ASSIGN_BUTTON = re.compile(r"^(교육 할당|Assign training)$", re.IGNORECASE)
_ASSIGN_TO_USER_ITEM = re.compile(r"^(사용자에 할당|Assign to users?)$", re.IGNORECASE)
_USER_DIALOG = re.compile(r"^(사용자 선택|Select users?)$", re.IGNORECASE)
_USER_SEARCH = re.compile(r"^(사용자 찾기|Find users?)$", re.IGNORECASE)
_USER_TABLE = re.compile(r"^(사용자|Users?)$", re.IGNORECASE)
_ASSIGN_BUTTON = re.compile(r"^(할당|Assign)$", re.IGNORECASE)
_USER_GRID = re.compile(r"^(사용자|Users?)$", re.IGNORECASE)


class AwsAssignmentError(RuntimeError):
    """Base error for a safe, stopped AWS assignment attempt."""


class AwsAssignmentCancelled(AwsAssignmentError):
    """Raised when Pause/Stop control cancels the current operation."""


class AwsUserLookupError(AwsAssignmentError):
    """Raised unless the search produces exactly one matching user row."""


@dataclass(frozen=True, slots=True)
class AwsAssignmentResult:
    """Verified result without exposing the searched value to application logs."""

    matched_user_label: str


class AwsSkillBuilderAssignment:
    """Assign one user on the observed AWS Skill Builder training-detail page."""

    def __init__(
        self,
        page: Page,
        stop_event: threading.Event,
        pause_event: threading.Event,
        on_stage: Callable[[str, str], None] | None = None,
        *,
        lookup_timeout_ms: int = 15_000,
        result_stable_ms: int = 500,
        invalid_result_stable_ms: int = 3_000,
        verification_timeout_ms: int = 20_000,
        enforce_aws_host: bool = True,
    ) -> None:
        self._page = page
        self._stop_event = stop_event
        self._pause_event = pause_event
        self._on_stage = on_stage or (lambda _stage, _message: None)
        self._lookup_timeout_ms = max(1_000, lookup_timeout_ms)
        self._result_stable_ms = max(300, result_stable_ms)
        self._invalid_result_stable_ms = max(
            self._result_stable_ms,
            invalid_result_stable_ms,
        )
        self._verification_timeout_ms = max(1_000, verification_timeout_ms)
        self._enforce_aws_host = enforce_aws_host

    def validate_ready(self, expected_url: str | None = None) -> str:
        """Verify the intended training page without clicking or entering data."""
        self._validate_target_page(expected_url)
        self._checkpoint()
        assign_button = self._single(
            self._page.get_by_role(
                "button",
                name=_TRAINING_ASSIGN_BUTTON,
                exact=True,
            ),
            "교육 할당 버튼",
        )
        if not assign_button.is_visible():
            raise AwsAssignmentError(
                "교육 할당 버튼이 현재 화면에 보이지 않습니다. "
                "로그인 후 교육 상세 페이지가 완전히 열린 상태인지 확인하세요."
            )
        if not assign_button.is_enabled():
            raise AwsAssignmentError("교육 할당 버튼이 아직 활성화되지 않았습니다.")
        return self._page.url

    def assign_user(
        self,
        search_value: str,
        expected_url: str | None = None,
    ) -> AwsAssignmentResult:
        """Run menu -> user search -> single row -> checkbox -> assign -> verify."""
        query = search_value.strip()
        if not query:
            raise AwsUserLookupError("빈 Excel 값은 사용자 검색에 사용할 수 없습니다.")
        self._validate_target_page(expected_url)
        self._checkpoint()

        self._stage("OPENING_ASSIGNMENT", "교육 할당 메뉴를 여는 중입니다.")
        dialog = self._open_user_dialog()
        self._checkpoint()

        self._stage("SEARCHING_USER", "사용자 찾기에 현재 셀 값을 입력했습니다.")
        search = self._single(
            dialog.get_by_role("combobox", name=_USER_SEARCH, exact=True),
            "사용자 찾기 입력 영역",
        )
        table = self._single(
            dialog.get_by_role("table", name=_USER_TABLE, exact=True),
            "사용자 검색 결과 표",
        )
        baseline_signature = self._row_signature(self._data_rows(table))
        search.fill(query, timeout=10_000)
        if search.input_value(timeout=5_000).strip() != query:
            raise AwsAssignmentError("사용자 찾기 입력값을 확인하지 못했습니다.")

        row = self._wait_for_one_matching_row(table, query, baseline_signature)
        self._stage("SELECTING_USER", "검색 결과 1개를 확인하고 선택했습니다.")
        checkbox = self._single(row.get_by_role("checkbox"), "검색 결과 체크박스")
        checkbox.check(timeout=5_000)
        if not checkbox.is_checked():
            raise AwsAssignmentError("검색 결과 체크박스가 선택되지 않았습니다.")
        matched_label = self._matched_user_label(row, query)

        assign_button = dialog.get_by_test_id(
            "assign-classroom-training-assign-modal-btn"
        )
        if assign_button.count() != 1:
            assign_button = self._single(
                dialog.get_by_role("button", name=_ASSIGN_BUTTON, exact=True),
                "할당 버튼",
            )
        self._wait_until_enabled(assign_button, 5_000)
        self._checkpoint()

        self._stage("ASSIGNING_USER", "선택한 사용자를 교육에 할당하는 중입니다.")
        assign_button.click(timeout=10_000)
        self._wait_until_hidden(dialog, 15_000)

        self._stage("VERIFYING_ASSIGNMENT", "할당 결과를 확인하는 중입니다.")
        self._wait_until_user_appears(matched_label)
        self._checkpoint()
        return AwsAssignmentResult(matched_user_label=matched_label)

    def _validate_target_page(self, expected_url: str | None = None) -> None:
        if not self._enforce_aws_host:
            return
        parsed = urlparse(self._page.url)
        valid_host = parsed.hostname == "skillbuilder.aws"
        valid_path = "/admin/organization/modality/curriculum/training/" in parsed.path
        if not (valid_host and valid_path):
            raise AwsAssignmentError(
                "AWS Skill Builder 교육 상세 페이지가 아닙니다. URL을 확인하세요."
            )
        if expected_url and not self._same_training_destination(
            expected_url,
            self._page.url,
        ):
            raise AwsAssignmentError(
                "현재 브라우저가 입력한 교육 상세 URL과 다른 페이지에 있습니다. "
                "대상 교육과 조직을 확인하세요."
            )

    def _open_user_dialog(self) -> Locator:
        dialog = self._page.get_by_role("dialog", name=_USER_DIALOG, exact=True)
        if dialog.count() == 1 and dialog.is_visible():
            return dialog

        assign_menu_button = self._single(
            self._page.get_by_role(
                "button",
                name=_TRAINING_ASSIGN_BUTTON,
                exact=True,
            ),
            "교육 할당 버튼",
        )
        assign_menu_button.click(timeout=10_000)
        self._checkpoint()
        assign_to_user = self._single(
            self._page.get_by_role(
                "menuitem",
                name=_ASSIGN_TO_USER_ITEM,
                exact=True,
            ),
            "사용자에 할당 메뉴",
        )
        assign_to_user.click(timeout=10_000)
        self._wait_until_visible(dialog, 10_000)
        return self._single(dialog, "사용자 선택 창")

    def _wait_for_one_matching_row(
        self,
        table: Locator,
        query: str,
        baseline_signature: tuple[str, ...],
    ) -> Locator:
        deadline = time.monotonic() + self._lookup_timeout_ms / 1000
        normalized_query = self._normalize(query)
        last_signature: tuple[str, ...] | None = None
        stable_since = time.monotonic()
        observed_change = False

        while time.monotonic() < deadline:
            self._checkpoint()
            rows = self._data_rows(table)
            signature = self._row_signature(rows)
            if signature != baseline_signature:
                observed_change = True
            if signature != last_signature:
                last_signature = signature
                stable_since = time.monotonic()

            stable_ms = (time.monotonic() - stable_since) * 1000
            is_one_match = len(signature) == 1 and normalized_query in self._normalize(
                signature[0]
            )
            if is_one_match and stable_ms >= self._result_stable_ms:
                return rows.first
            if (
                observed_change
                and not is_one_match
                and stable_ms >= self._invalid_result_stable_ms
            ):
                raise AwsUserLookupError(self._lookup_error_message(len(signature)))
            self._interruptible_wait(0.1)

        final_count = self._data_rows(table).count()
        raise AwsUserLookupError(self._lookup_error_message(final_count))

    def _wait_until_user_appears(self, matched_label: str) -> None:
        deadline = time.monotonic() + self._verification_timeout_ms / 1000
        while time.monotonic() < deadline:
            self._checkpoint()
            grids = self._page.get_by_role("grid", name=_USER_GRID, exact=True)
            if grids.count() == 1:
                matches = grids.get_by_text(matched_label, exact=True)
                if matches.count() >= 1 and matches.first.is_visible():
                    return
            self._interruptible_wait(0.15)
        raise AwsAssignmentError(
            "할당 창은 닫혔지만 사용자 목록에서 결과를 확인하지 못했습니다. "
            "현재 셀을 이동하지 않았으므로 페이지에서 할당 여부를 확인하세요."
        )

    def _wait_until_enabled(self, locator: Locator, timeout_ms: int) -> None:
        deadline = time.monotonic() + timeout_ms / 1000
        while time.monotonic() < deadline:
            self._checkpoint()
            if locator.is_enabled():
                return
            self._interruptible_wait(0.1)
        raise AwsAssignmentError("할당 버튼이 활성화되지 않았습니다.")

    def _wait_until_visible(self, locator: Locator, timeout_ms: int) -> None:
        deadline = time.monotonic() + timeout_ms / 1000
        while time.monotonic() < deadline:
            self._checkpoint()
            if locator.count() == 1 and locator.is_visible():
                return
            self._interruptible_wait(0.1)
        raise AwsAssignmentError("사용자 선택 창이 열리지 않았습니다.")

    def _wait_until_hidden(self, locator: Locator, timeout_ms: int) -> None:
        deadline = time.monotonic() + timeout_ms / 1000
        while time.monotonic() < deadline:
            self._checkpoint()
            if locator.count() == 0 or not locator.is_visible():
                return
            self._interruptible_wait(0.1)
        raise AwsAssignmentError("할당 후 사용자 선택 창이 닫히지 않았습니다.")

    def _checkpoint(self) -> None:
        while self._pause_event.is_set():
            if self._stop_event.wait(0.1):
                raise AwsAssignmentCancelled("AWS 사용자 할당이 중지되었습니다.")
        if self._stop_event.is_set():
            raise AwsAssignmentCancelled("AWS 사용자 할당이 중지되었습니다.")
        if self._page.is_closed():
            raise AwsAssignmentError("브라우저 페이지가 닫혔습니다.")

    def _interruptible_wait(self, seconds: float) -> None:
        if self._stop_event.wait(seconds):
            raise AwsAssignmentCancelled("AWS 사용자 할당이 중지되었습니다.")

    def _stage(self, stage: str, message: str) -> None:
        self._on_stage(stage, message)

    @staticmethod
    def _data_rows(table: Locator) -> Locator:
        return table.locator('tbody tr:has(input[type="checkbox"])')

    @staticmethod
    def _row_signature(rows: Locator) -> tuple[str, ...]:
        return tuple(text.strip() for text in rows.all_inner_texts())

    @staticmethod
    def _normalize(value: str) -> str:
        return " ".join(value.casefold().split())

    @staticmethod
    def _same_training_destination(expected_url: str, actual_url: str) -> bool:
        expected = urlparse(expected_url)
        actual = urlparse(actual_url)
        expected_org = parse_qs(expected.query).get("orgId", [])
        actual_org = parse_qs(actual.query).get("orgId", [])
        return (
            expected.scheme.casefold() == actual.scheme.casefold()
            and (expected.hostname or "").casefold()
            == (actual.hostname or "").casefold()
            and expected.path.rstrip("/") == actual.path.rstrip("/")
            and expected_org == actual_org
        )

    @staticmethod
    def _lookup_error_message(count: int) -> str:
        return (
            f"사용자 검색 결과가 {count}개입니다. 정확히 1개일 때만 할당합니다. "
            "현재 Excel 셀은 이동하지 않았습니다."
        )

    @staticmethod
    def _matched_user_label(row: Locator, query: str) -> str:
        links = [text.strip() for text in row.get_by_role("link").all_inner_texts()]
        for label in links:
            if label:
                return label
        return query

    @staticmethod
    def _single(locator: Locator, description: str) -> Locator:
        count = locator.count()
        if count != 1:
            raise AwsAssignmentError(
                f"{description}이(가) 정확히 1개여야 합니다. 현재 {count}개입니다."
            )
        return locator
